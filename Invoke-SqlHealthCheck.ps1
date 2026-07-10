<#
=====================================================================================
 SQL DBA Champs  |  Centralized SQL Server Health Check
-------------------------------------------------------------------------------------
 Invoke-SqlHealthCheck.ps1

 Purpose : Connect to many on-prem SQL Server instances from ONE central box,
           run a fixed set of health checks against each, roll the results up
           into a single HTML report + a single Excel workbook, and send ONE
           consolidated email covering every server.

 Method  : Central-orchestrator pattern (recommended over per-server Agent jobs
           when you want a single consolidated report/email). Schedule this
           script with a SQL Agent CmdExec job or Task Scheduler on the monitor box.

 Design  : Pure ADO.NET (System.Data.SqlClient) for the SQL work - NO SqlServer /
           dbatools module required. Excel export uses the ImportExcel module if
           present and falls back to CSV otherwise. Reachability of the SQL service
           is inferred from the connection test; an optional CIM check can verify
           the Windows service state directly.

 Auth    : Windows auth by default. Use -SqlAuth with -SqlUser / -SqlPassword for
           SQL logins. The account needs VIEW SERVER STATE + read on msdb on each
           target instance.

 Author  : PMSQLDBA  |  sqldbachamps.com
 Tested  : Windows PowerShell 5.1 (and PowerShell 7) against SQL Server 2016-2022
=====================================================================================
#>

[CmdletBinding()]
param(
    # Path to a plain-text list of instances (one per line, "#" comments allowed).
    # If omitted, the $DefaultInstances array below is used.
    [string]   $ConfigPath   = "$PSScriptRoot\servers.txt",

    # Where the HTML / Excel report files are written.
    [string]   $OutputPath   = "$PSScriptRoot\Reports",

    # --- Email settings ---
    [string]   $SmtpServer   = 'smtp.yourcompany.local',
    [int]      $SmtpPort     = 25,
    [string]   $MailFrom     = 'sql-healthcheck@yourcompany.local',
    [string[]] $MailTo       = @('dba-team@yourcompany.local'),
    [switch]   $UseSsl,
    [switch]   $NoEmail,                 # generate reports but skip sending

    # --- Authentication ---
    [switch]   $SqlAuth,
    [string]   $SqlUser,
    [string]   $SqlPassword,

    # --- Optional Windows service verification (needs CIM/WMI remoting) ---
    [switch]   $CheckWindowsService
)

# -------------------------------------------------------------------------------------
# 0. Default instance list  (replace with your estate, or use servers.txt)
# -------------------------------------------------------------------------------------
$DefaultInstances = @(
    'VMSQL01'
    'VMSQL02'
    'VMSQL03'
    # 'PRODSQL01\FINANCE'
    # 'PRODSQL02,1433'
)

# -------------------------------------------------------------------------------------
# 1. Thresholds  (tune to taste - drive every OK / WARNING / CRITICAL decision)
# -------------------------------------------------------------------------------------
$Thresholds = @{
    DiskFreePctWarn       = 15      # % free -> WARNING at/below
    DiskFreePctCrit       = 8       # % free -> CRITICAL at/below
    FullBackupWarnHours   = 26      # last FULL older than this -> WARNING
    FullBackupCritHours   = 50      # last FULL older than this -> CRITICAL
    LogBackupWarnHours    = 4        # last LOG (FULL/BULK recovery) older -> WARNING
    LogBackupCritHours    = 12       # last LOG older -> CRITICAL
    FailedJobWindowHours  = 24       # look back this far for failed Agent jobs
    AgentStoppedIsCritical = $false  # $true -> a stopped (Automatic) SQL Agent grades CRITICAL, else WARNING
    LongRunSeconds        = 300      # ACTIVE requests running longer than this are reported
    LongRunCritSeconds    = 1800     # long-running requests past this are CRITICAL
    LongTxnSeconds        = 300      # OPEN transactions older than this are reported (even if idle)
    LongTxnCritSeconds    = 1800     # open transactions past this are CRITICAL
    BlockingWarnSeconds   = 30       # any blocking >= this -> WARNING
    BlockingCritSeconds   = 120      # blocking >= this -> CRITICAL
    CpuWarnPct            = 80       # SQL process CPU% -> WARNING
    CpuCritPct            = 92       # SQL process CPU% -> CRITICAL
    MemAvailWarnMB        = 2048     # OS available memory below this -> WARNING
    MemAvailCritMB        = 512      # OS available memory below this -> CRITICAL
    ErrorLogWindowHours   = 24       # scan this much of the current error log
    ConnectTimeoutSec     = 10
    QueryTimeoutSec       = 45
}

# Error-log patterns. xp_readerrorlog does a substring search per term, so we run it
# once per term and merge. These catch login failures ("Login failed for user..."),
# numbered errors, severity 16-25, and other real problems. Tune to your noise profile.
$ErrorLogSearchTerms   = @('Login failed','Error:','Severity: 1','Severity: 2','failed','corrupt','deadlock','cannot')
# Any captured line matching one of these escalates the whole check to CRITICAL.
$ErrorLogCriticalTerms = @('Severity: 2','corrupt','consistency','I/O error','stack dump','has encountered','823','824','825')

# House palette (SQL DBA Champs)
$Brand = @{
    Navy    = '#1F3864'
    Accent  = '#2E75B6'
    OkBg     = '#E8F5E9'; OkFg     = '#1B5E20'
    WarnBg   = '#FFF4E5'; WarnFg   = '#8A5300'
    CritBg   = '#FDECEA'; CritFg   = '#B71C1C'
    HeadRow  = '#EAF1F8'
    Border   = '#C9D3E0'
    Muted    = '#6B7280'
}

# =====================================================================================
# 2. Helpers
# =====================================================================================

function Get-InstanceList {
    param([string]$Path, [string[]]$Fallback)
    if (Test-Path -LiteralPath $Path) {
        return Get-Content -LiteralPath $Path |
            ForEach-Object { $_.Trim() } |
            Where-Object { $_ -and -not $_.StartsWith('#') }
    }
    return $Fallback
}

function Invoke-HcQuery {
    <#
      Runs a query with System.Data.SqlClient and returns an array of PSCustomObjects.
      Throws on connection/query failure so the caller can flag the server DOWN.
    #>
    param(
        [Parameter(Mandatory)] [string] $Instance,
        [Parameter(Mandatory)] [string] $Query,
        [string] $Database = 'master'
    )
    Add-Type -AssemblyName System.Data | Out-Null

    if ($SqlAuth) {
        $auth = "User ID=$SqlUser;Password=$SqlPassword"
    } else {
        $auth = 'Integrated Security=SSPI'
    }
    $connStr = "Server=$Instance;Database=$Database;$auth;" +
               "Connect Timeout=$($Thresholds.ConnectTimeoutSec);Application Name=SqlHealthCheck;TrustServerCertificate=True"

    $conn = New-Object System.Data.SqlClient.SqlConnection $connStr
    $conn.Open()
    try {
        $cmd = $conn.CreateCommand()
        $cmd.CommandText    = $Query
        $cmd.CommandTimeout = $Thresholds.QueryTimeoutSec
        $adapter = New-Object System.Data.SqlClient.SqlDataAdapter $cmd
        $ds = New-Object System.Data.DataSet
        [void]$adapter.Fill($ds)
        if ($ds.Tables.Count -eq 0) { return @() }
        # Project DataRows into clean PSCustomObjects (column values only, DBNull -> $null)
        # so downstream rendering never sees DataRow base members (RowState, ItemArray, ...).
        $table = $ds.Tables[0]
        $cols  = @($table.Columns | ForEach-Object { $_.ColumnName })
        foreach ($row in $table.Rows) {
            $o = [ordered]@{}
            foreach ($c in $cols) {
                $v = $row[$c]
                if ($v -is [System.DBNull]) { $v = $null }
                $o[$c] = $v
            }
            [pscustomobject]$o
        }
    }
    finally { $conn.Close(); $conn.Dispose() }
}

function New-Check {
    param([string]$Name, [string]$Status, [string]$Summary, $Rows = @())
    [pscustomobject]@{
        Name    = $Name
        Status  = $Status        # OK | WARNING | CRITICAL | ERROR
        Summary = $Summary
        Rows    = @($Rows)
    }
}

# Worst-status roll-up (ERROR is treated as CRITICAL for colouring)
function Get-WorstStatus {
    param([string[]]$Statuses)
    if ($Statuses -contains 'CRITICAL' -or $Statuses -contains 'ERROR') { return 'CRITICAL' }
    if ($Statuses -contains 'WARNING') { return 'WARNING' }
    return 'OK'
}

# =====================================================================================
# 3. Check queries
# =====================================================================================

$Q = @{}

# Engine, Agent, Browser, Full-text service states from inside SQL (needs VIEW SERVER STATE only).
$Q.Services = @"
SELECT servicename, status_desc, startup_type_desc
FROM sys.dm_server_services;
"@

$Q.DiskSpace = @"
SELECT DISTINCT
    vs.volume_mount_point                                             AS Drive,
    CAST(vs.total_bytes     / 1073741824.0 AS DECIMAL(18,2))          AS TotalGB,
    CAST(vs.available_bytes / 1073741824.0 AS DECIMAL(18,2))          AS FreeGB,
    CAST(vs.available_bytes * 100.0 / NULLIF(vs.total_bytes,0) AS DECIMAL(5,2)) AS FreePct
FROM sys.master_files mf
CROSS APPLY sys.dm_os_volume_stats(mf.database_id, mf.file_id) vs
ORDER BY FreePct;
"@

$Q.DbStatus = @"
SELECT name AS DatabaseName, state_desc AS State,
       user_access_desc AS UserAccess, recovery_model_desc AS RecoveryModel
FROM sys.databases
ORDER BY CASE WHEN state_desc <> 'ONLINE' THEN 0 ELSE 1 END, name;
"@

$Q.Backups = @"
SELECT d.name AS DatabaseName, d.recovery_model_desc AS RecoveryModel,
       MAX(CASE WHEN b.type='D' THEN b.backup_finish_date END) AS LastFull,
       MAX(CASE WHEN b.type='I' THEN b.backup_finish_date END) AS LastDiff,
       MAX(CASE WHEN b.type='L' THEN b.backup_finish_date END) AS LastLog
FROM sys.databases d
LEFT JOIN msdb.dbo.backupset b ON b.database_name = d.name
WHERE d.name <> 'tempdb' AND d.state_desc = 'ONLINE'
GROUP BY d.name, d.recovery_model_desc
ORDER BY d.name;
"@

$Q.FailedJobs = @"
SELECT j.name AS JobName,
       msdb.dbo.agent_datetime(h.run_date, h.run_time) AS LastRunTime,
       h.run_status AS RunStatus,
       LEFT(h.message, 300) AS Message
FROM msdb.dbo.sysjobs j
JOIN msdb.dbo.sysjobhistory h ON h.job_id = j.job_id AND h.step_id = 0
WHERE h.run_status = 0
  AND msdb.dbo.agent_datetime(h.run_date, h.run_time) >= DATEADD(HOUR, -$($Thresholds.FailedJobWindowHours), GETDATE())
ORDER BY LastRunTime DESC;
"@

$Q.Blocking = @"
SELECT r.session_id            AS BlockedSpid,
       r.blocking_session_id   AS BlockingSpid,
       DB_NAME(r.database_id)  AS DatabaseName,
       r.wait_type             AS WaitType,
       CAST(r.wait_time/1000.0 AS DECIMAL(10,1)) AS WaitSeconds,
       r.command               AS Command,
       LEFT(t.text,200)        AS QueryText
FROM sys.dm_exec_requests r
OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) t
WHERE r.blocking_session_id <> 0
ORDER BY r.wait_time DESC;
"@

$Q.LongRunning = @"
SELECT r.session_id AS Spid, s.login_name AS LoginName,
       DB_NAME(r.database_id) AS DatabaseName, r.status AS Status, r.command AS Command,
       CAST(r.total_elapsed_time/1000.0 AS DECIMAL(10,1)) AS ElapsedSeconds,
       LEFT(t.text,200) AS QueryText
FROM sys.dm_exec_requests r
JOIN sys.dm_exec_sessions s ON s.session_id = r.session_id
OUTER APPLY sys.dm_exec_sql_text(r.sql_handle) t
WHERE r.session_id <> @@SPID AND s.is_user_process = 1
  AND r.total_elapsed_time/1000.0 >= $($Thresholds.LongRunSeconds)
ORDER BY r.total_elapsed_time DESC;
"@

# Long-running OPEN TRANSACTIONS. Unlike dm_exec_requests (active queries only), the
# transaction DMVs surface a transaction that is open but whose session is idle/sleeping
# (BEGIN TRAN with no COMMIT) - the classic lock-holder / log-pinner DBCC OPENTRAN finds.
$Q.LongTxn = @"
SELECT
    s.session_id                                           AS Spid,
    s.login_name                                           AS LoginName,
    DB_NAME(dbt.database_id)                               AS DatabaseName,
    DATEDIFF(SECOND, at.transaction_begin_time, GETDATE()) AS DurationSeconds,
    CONVERT(VARCHAR(19), at.transaction_begin_time, 120)   AS BeginTime,
    s.status                                               AS SessionStatus,
    LEFT(ISNULL(s.host_name,''),30)                        AS HostName,
    LEFT(ISNULL(s.program_name,''),50)                     AS Program,
    LEFT(ISNULL(txt.text,''),200)                          AS LastQuery
FROM sys.dm_tran_active_transactions at
JOIN sys.dm_tran_session_transactions st ON st.transaction_id = at.transaction_id
JOIN sys.dm_exec_sessions s              ON s.session_id = st.session_id
OUTER APPLY (SELECT TOP 1 dt.database_id FROM sys.dm_tran_database_transactions dt
             WHERE dt.transaction_id = at.transaction_id ORDER BY dt.database_id) dbt
LEFT JOIN sys.dm_exec_connections c      ON c.session_id = s.session_id
OUTER APPLY sys.dm_exec_sql_text(c.most_recent_sql_handle) txt
WHERE st.is_user_transaction = 1
  AND s.is_user_process = 1
  AND s.session_id <> @@SPID
  AND at.transaction_begin_time <= DATEADD(SECOND, -$($Thresholds.LongTxnSeconds), GETDATE())
ORDER BY at.transaction_begin_time ASC;
"@

$Q.Cpu = @"
SELECT TOP 1 SQLProcessUtilization, SystemIdle,
       100 - SystemIdle - SQLProcessUtilization AS OtherProcessUtilization
FROM (
  SELECT record.value('(./Record/@id)[1]','int') AS record_id,
         record.value('(./Record/SchedulerMonitorEvent/SystemHealth/SystemIdle)[1]','int') AS SystemIdle,
         record.value('(./Record/SchedulerMonitorEvent/SystemHealth/ProcessUtilization)[1]','int') AS SQLProcessUtilization
  FROM (
    SELECT CONVERT(xml, record) AS record
    FROM sys.dm_os_ring_buffers
    WHERE ring_buffer_type = N'RING_BUFFER_SCHEDULER_MONITOR'
      AND record LIKE '%<SystemHealth>%'
  ) x
) y
ORDER BY record_id DESC;
"@

$Q.Memory = @"
SELECT
    CAST((SELECT physical_memory_in_use_kb FROM sys.dm_os_process_memory)/1024.0 AS DECIMAL(18,0)) AS SqlMemoryUsedMB,
    CAST(total_physical_memory_kb/1024.0     AS DECIMAL(18,0)) AS TotalServerMemoryMB,
    CAST(available_physical_memory_kb/1024.0 AS DECIMAL(18,0)) AS AvailableServerMemoryMB,
    system_memory_state_desc AS MemoryState
FROM sys.dm_os_sys_memory;
"@

# Build one INSERT..EXEC per search term (xp_readerrorlog can't OR terms in a single call).
$elInserts = ($ErrorLogSearchTerms | ForEach-Object {
    "INSERT INTO #el EXEC sys.xp_readerrorlog 0, 1, N'$($_.Replace("'","''"))', NULL, @start, NULL, N'desc';"
}) -join "`n"

# NOTE: xp_readerrorlog requires securityadmin membership (or GRANT EXEC on it).
$Q.ErrorLog = @"
IF OBJECT_ID('tempdb..#el') IS NOT NULL DROP TABLE #el;
CREATE TABLE #el (LogDate DATETIME, ProcessInfo NVARCHAR(100), [Text] NVARCHAR(MAX));
DECLARE @start DATETIME = DATEADD(HOUR, -$($Thresholds.ErrorLogWindowHours), GETDATE());
$elInserts
SELECT DISTINCT TOP 60
    CONVERT(VARCHAR(19), LogDate, 120) AS LogTime,
    LTRIM(RTRIM(ProcessInfo))         AS Source,
    LEFT(LTRIM([Text]), 400)          AS ErrorMessage
FROM #el
WHERE [Text] NOT LIKE '%Login succeeded%'
  AND [Text] NOT LIKE '%error log has been reinitialized%'
  AND [Text] NOT LIKE '%This instance of SQL Server has been using a process ID%'
ORDER BY LogTime DESC;
DROP TABLE #el;
"@

# =====================================================================================
# 4. Per-server evaluation
# =====================================================================================

function Test-ServerHealth {
    param([string]$Instance)

    $server = [pscustomobject]@{
        Instance     = $Instance
        Reachable    = $true
        Overall      = 'OK'
        Checks       = @()
        Error        = $null
        WindowsSvc   = $null
        AgentRunning = $null
    }

    # --- Optional Windows service state (needs remoting) ---
    if ($CheckWindowsService) {
        try {
            $host2 = ($Instance -split '[\\,]')[0]
            $svc = Get-CimInstance -ComputerName $host2 -ClassName Win32_Service `
                    -Filter "Name LIKE 'MSSQL%' AND Name NOT LIKE '%CEIP%'" -ErrorAction Stop |
                   Select-Object -First 1
            if ($svc) { $server.WindowsSvc = "$($svc.Name) = $($svc.State)" }
        } catch { $server.WindowsSvc = 'service query failed' }
    }

    # --- Reachability = service status ---
    try {
        [void](Invoke-HcQuery -Instance $Instance -Query 'SELECT 1 AS ok;')
    } catch {
        $server.Reachable = $false
        $server.Overall   = 'CRITICAL'
        $server.Error     = $_.Exception.Message
        $server.Checks   += New-Check 'SQL Service' 'CRITICAL' 'Cannot connect - service down or unreachable'
        return $server
    }
    # --- SQL Service: engine reachable + service inventory (engine, agent, browser, ...) ---
    $svcRows = @()
    try { $svcRows = @(Invoke-HcQuery -Instance $Instance -Query $Q.Services) } catch { }
    $svcInv = @($svcRows | ForEach-Object {
        [pscustomobject]@{ Service = $_.servicename; Status = $_.status_desc; StartupType = $_.startup_type_desc }
    })
    $server.Checks += New-Check 'SQL Service' 'OK' 'Database Engine running' $svcInv

    # --- SQL Agent service state (the gap: engine can be up while Agent is stopped) ---
    $agent = $svcRows | Where-Object { $_.servicename -like 'SQL Server Agent*' } | Select-Object -First 1
    if ($agent) {
        $running = $agent.status_desc -eq 'Running'
        $server.AgentRunning = $running
        if ($running) {
            $server.Checks += New-Check 'SQL Agent' 'OK' "$($agent.servicename) running"
        }
        elseif ($agent.startup_type_desc -eq 'Automatic') {
            $sev = if ($Thresholds.AgentStoppedIsCritical) { 'CRITICAL' } else { 'WARNING' }
            $server.Checks += New-Check 'SQL Agent' $sev "$($agent.servicename) is STOPPED (startup: Automatic) - scheduled jobs are NOT running"
        }
        else {
            # stopped but Manual/Disabled -> almost certainly intentional, don't cry wolf
            $server.Checks += New-Check 'SQL Agent' 'OK' "$($agent.servicename) stopped (startup: $($agent.startup_type_desc), not auto-start)"
        }
    } else {
        $server.AgentRunning = $null
        $server.Checks += New-Check 'SQL Agent' 'OK' 'Agent state unavailable from dm_server_services'
    }

    # --- Disk space ---
    try {
        $rows = @(Invoke-HcQuery -Instance $Instance -Query $Q.DiskSpace)
        $crit = $rows | Where-Object { [double]$_.FreePct -le $Thresholds.DiskFreePctCrit }
        $warn = $rows | Where-Object { [double]$_.FreePct -le $Thresholds.DiskFreePctWarn -and [double]$_.FreePct -gt $Thresholds.DiskFreePctCrit }
        $st   = if ($crit) {'CRITICAL'} elseif ($warn) {'WARNING'} else {'OK'}
        $sum  = if ($crit -or $warn) { "$(@($crit).Count + @($warn).Count) volume(s) low on space" } else { "$(@($rows).Count) volume(s) healthy" }
        $server.Checks += New-Check 'Disk Space' $st $sum $rows
    } catch { $server.Checks += New-Check 'Disk Space' 'ERROR' $_.Exception.Message }

    # --- Database status ---
    try {
        $rows = @(Invoke-HcQuery -Instance $Instance -Query $Q.DbStatus)
        $bad  = $rows | Where-Object { $_.State -in @('SUSPECT','RECOVERY_PENDING','EMERGENCY') }
        $off  = $rows | Where-Object { $_.State -eq 'OFFLINE' }
        $rest = $rows | Where-Object { $_.State -notin @('ONLINE','OFFLINE') } | Where-Object { $_.State -notin @('SUSPECT','RECOVERY_PENDING','EMERGENCY') }
        $st   = if ($bad) {'CRITICAL'} elseif ($off -or $rest) {'WARNING'} else {'OK'}
        $issue = @($rows | Where-Object { $_.State -ne 'ONLINE' })
        $sum  = if ($issue.Count) { "$($issue.Count) database(s) not ONLINE" } else { "All $(@($rows).Count) databases ONLINE" }
        # only surface the non-online rows in detail to keep the report tight
        $detail = if ($issue.Count) { $issue } else { @() }
        $server.Checks += New-Check 'Database Status' $st $sum $detail
    } catch { $server.Checks += New-Check 'Database Status' 'ERROR' $_.Exception.Message }

    # --- Backups ---
    try {
        $rows = @(Invoke-HcQuery -Instance $Instance -Query $Q.Backups)
        $now  = Get-Date
        $flag = foreach ($r in $rows) {
            $ageFull = if ($r.LastFull -is [datetime]) { ($now - $r.LastFull).TotalHours } else { [double]::MaxValue }
            $ageLog  = if ($r.LastLog  -is [datetime]) { ($now - $r.LastLog ).TotalHours } else { [double]::MaxValue }
            $isSystem = $r.DatabaseName -in @('master','model','msdb')
            $needLog  = ($r.RecoveryModel -in @('FULL','BULK_LOGGED')) -and -not $isSystem
            $s = 'OK'
            if     ($ageFull -ge $Thresholds.FullBackupCritHours) { $s = 'CRITICAL' }
            elseif ($ageFull -ge $Thresholds.FullBackupWarnHours) { $s = 'WARNING' }
            if ($needLog) {
                if     ($ageLog -ge $Thresholds.LogBackupCritHours) { $s = 'CRITICAL' }
                elseif ($ageLog -ge $Thresholds.LogBackupWarnHours -and $s -eq 'OK') { $s = 'WARNING' }
            }
            if ($s -ne 'OK') {
                [pscustomobject]@{
                    DatabaseName  = $r.DatabaseName
                    RecoveryModel = $r.RecoveryModel
                    LastFull      = if ($r.LastFull -is [datetime]) { $r.LastFull.ToString('yyyy-MM-dd HH:mm') } else { 'NEVER' }
                    LastLog       = if ($r.LastLog  -is [datetime]) { $r.LastLog.ToString('yyyy-MM-dd HH:mm') }  else { 'n/a' }
                    Issue         = $s
                }
            }
        }
        $flag = @($flag)
        $st   = Get-WorstStatus ($flag.Issue)
        $sum  = if ($flag.Count) { "$($flag.Count) database(s) with stale/missing backups" } else { "All backups current ($(@($rows).Count) db)" }
        $server.Checks += New-Check 'Backup Status' $st $sum $flag
    } catch { $server.Checks += New-Check 'Backup Status' 'ERROR' $_.Exception.Message }

    # --- Failed Agent jobs ---
    try {
        if ($server.AgentRunning -eq $false) {
            # Agent is down: "no failed jobs" would be a false OK because nothing is running.
            $server.Checks += New-Check 'Failed Agent Jobs' 'WARNING' 'SQL Server Agent is stopped - no jobs are running to succeed or fail'
        } else {
            $rows = @(Invoke-HcQuery -Instance $Instance -Query $Q.FailedJobs)
            $st   = if ($rows.Count) {'CRITICAL'} else {'OK'}
            $sum  = if ($rows.Count) { "$($rows.Count) failed job run(s) in last $($Thresholds.FailedJobWindowHours)h" } else { "No failed jobs in last $($Thresholds.FailedJobWindowHours)h" }
            $server.Checks += New-Check 'Failed Agent Jobs' $st $sum $rows
        }
    } catch { $server.Checks += New-Check 'Failed Agent Jobs' 'ERROR' $_.Exception.Message }

    # --- Blocking ---
    try {
        $rows = @(Invoke-HcQuery -Instance $Instance -Query $Q.Blocking)
        $maxW = ($rows | Measure-Object WaitSeconds -Maximum).Maximum
        $st   = if ($maxW -ge $Thresholds.BlockingCritSeconds) {'CRITICAL'} elseif ($rows.Count) {'WARNING'} else {'OK'}
        $sum  = if ($rows.Count) { "$($rows.Count) blocked session(s), max wait $([int]$maxW)s" } else { 'No blocking' }
        $server.Checks += New-Check 'Blocking Sessions' $st $sum $rows
    } catch { $server.Checks += New-Check 'Blocking Sessions' 'ERROR' $_.Exception.Message }

    # --- Long-running requests (active queries) ---
    try {
        $rows = @(Invoke-HcQuery -Instance $Instance -Query $Q.LongRunning)
        $maxE = ($rows | Measure-Object ElapsedSeconds -Maximum).Maximum
        $st   = if ($maxE -ge $Thresholds.LongRunCritSeconds) {'CRITICAL'} elseif ($rows.Count) {'WARNING'} else {'OK'}
        $sum  = if ($rows.Count) { "$($rows.Count) query(ies) > $($Thresholds.LongRunSeconds)s, max $([int]$maxE)s" } else { "None over $($Thresholds.LongRunSeconds)s" }
        $server.Checks += New-Check 'Long-Running Sessions' $st $sum $rows
    } catch { $server.Checks += New-Check 'Long-Running Sessions' 'ERROR' $_.Exception.Message }

    # --- Long-running open transactions (idle or active) ---
    try {
        $rows = @(Invoke-HcQuery -Instance $Instance -Query $Q.LongTxn)
        $maxD = ($rows | Measure-Object DurationSeconds -Maximum).Maximum
        $st   = if ($maxD -ge $Thresholds.LongTxnCritSeconds) {'CRITICAL'} elseif ($rows.Count) {'WARNING'} else {'OK'}
        $sum  = if ($rows.Count) {
            $mins = [math]::Round($maxD/60,1)
            "$($rows.Count) open transaction(s) > $($Thresholds.LongTxnSeconds)s, oldest $mins min"
        } else { "None open over $($Thresholds.LongTxnSeconds)s" }
        $server.Checks += New-Check 'Long-Running Transactions' $st $sum $rows
    } catch { $server.Checks += New-Check 'Long-Running Transactions' 'ERROR' $_.Exception.Message }

    # --- CPU ---
    try {
        $rows = @(Invoke-HcQuery -Instance $Instance -Query $Q.Cpu)
        $cpu  = if ($rows.Count) { [int]$rows[0].SQLProcessUtilization } else { $null }
        $st   = if ($cpu -ge $Thresholds.CpuCritPct) {'CRITICAL'} elseif ($cpu -ge $Thresholds.CpuWarnPct) {'WARNING'} else {'OK'}
        $sum  = if ($null -ne $cpu) { "SQL process CPU ~ $cpu% (system idle $($rows[0].SystemIdle)%)" } else { 'CPU sample unavailable' }
        $server.Checks += New-Check 'CPU Usage' $st $sum $rows
    } catch { $server.Checks += New-Check 'CPU Usage' 'ERROR' $_.Exception.Message }

    # --- Memory ---
    try {
        $rows = @(Invoke-HcQuery -Instance $Instance -Query $Q.Memory)
        $avail = if ($rows.Count) { [double]$rows[0].AvailableServerMemoryMB } else { $null }
        $state = if ($rows.Count) { $rows[0].MemoryState } else { '' }
        $st   = if ($avail -le $Thresholds.MemAvailCritMB) {'CRITICAL'} elseif ($avail -le $Thresholds.MemAvailWarnMB) {'WARNING'} else {'OK'}
        if ($state -match 'low|steady state.*low') { if ($st -eq 'OK') { $st = 'WARNING' } }
        $sum  = if ($null -ne $avail) { "OS available $([int]$avail) MB; SQL using $([int]$rows[0].SqlMemoryUsedMB) MB; state: $state" } else { 'Memory sample unavailable' }
        $server.Checks += New-Check 'Memory Usage' $st $sum $rows
    } catch { $server.Checks += New-Check 'Memory Usage' 'ERROR' $_.Exception.Message }

    # --- Error log ---
    try {
        $rows = @(Invoke-HcQuery -Instance $Instance -Query $Q.ErrorLog)
        $hasCrit = @($rows | Where-Object {
            $m = [string]$_.ErrorMessage
            ($ErrorLogCriticalTerms | Where-Object { $m -like "*$_*" }).Count -gt 0
        }).Count
        $logins  = @($rows | Where-Object { [string]$_.ErrorMessage -like '*Login failed*' }).Count
        $st  = if ($hasCrit) {'CRITICAL'} elseif ($rows.Count) {'WARNING'} else {'OK'}
        $sum = if ($rows.Count) {
            $extra = @()
            if ($hasCrit) { $extra += "$hasCrit severe" }
            if ($logins)  { $extra += "$logins login failure(s)" }
            $tail = if ($extra) { " ($($extra -join ', '))" } else { '' }
            "$($rows.Count) error entr(y/ies) in last $($Thresholds.ErrorLogWindowHours)h$tail"
        } else { "No errors in last $($Thresholds.ErrorLogWindowHours)h" }
        $server.Checks += New-Check 'Error Log' $st $sum $rows
    } catch { $server.Checks += New-Check 'Error Log' 'ERROR' $_.Exception.Message }

    $server.Overall = Get-WorstStatus ($server.Checks.Status)
    return $server
}

# =====================================================================================
# 5. HTML report
# =====================================================================================

function ConvertTo-Badge {
    param([string]$Status)
    switch ($Status) {
        'OK'       { "<span style='background:$($Brand.OkBg);color:$($Brand.OkFg);padding:2px 9px;border-radius:10px;font-weight:600;font-size:12px'>OK</span>" }
        'WARNING'  { "<span style='background:$($Brand.WarnBg);color:$($Brand.WarnFg);padding:2px 9px;border-radius:10px;font-weight:600;font-size:12px'>WARNING</span>" }
        'CRITICAL' { "<span style='background:$($Brand.CritBg);color:$($Brand.CritFg);padding:2px 9px;border-radius:10px;font-weight:600;font-size:12px'>CRITICAL</span>" }
        default    { "<span style='background:$($Brand.CritBg);color:$($Brand.CritFg);padding:2px 9px;border-radius:10px;font-weight:600;font-size:12px'>ERROR</span>" }
    }
}

function ConvertTo-DetailTable {
    param($Rows)
    $rows = @($Rows)
    if (-not $rows.Count) { return '' }
    $cols = $rows[0].PSObject.Properties.Name
    $th = ($cols | ForEach-Object { "<th style='text-align:left;padding:5px 9px;border-bottom:2px solid $($Brand.Border);font-size:12px;color:$($Brand.Navy)'>$_</th>" }) -join ''
    $tr = foreach ($r in $rows) {
        $td = ($cols | ForEach-Object {
            $v = $r.$_
            if ($v -is [datetime]) { $v = $v.ToString('yyyy-MM-dd HH:mm') }
            "<td style='padding:5px 9px;border-bottom:1px solid #EEF1F5;font-size:12px;font-family:Consolas,monospace'>$([System.Web.HttpUtility]::HtmlEncode([string]$v))</td>"
        }) -join ''
        "<tr>$td</tr>"
    }
    "<table style='border-collapse:collapse;width:100%;margin:6px 0 14px 0'><thead><tr>$th</tr></thead><tbody>$($tr -join '')</tbody></table>"
}

function New-HtmlReport {
    param($Servers, [datetime]$RunTime)
    Add-Type -AssemblyName System.Web | Out-Null

    $checkNames = @('SQL Service','SQL Agent','Disk Space','Database Status','Backup Status','Failed Agent Jobs',
                    'Blocking Sessions','Long-Running Sessions','Long-Running Transactions','CPU Usage','Memory Usage','Error Log')

    # ---- Summary matrix ----
    $matHead = "<th style='padding:8px 10px;text-align:left;background:$($Brand.Navy);color:#fff;font-size:12px'>Instance</th>" +
               "<th style='padding:8px 10px;background:$($Brand.Navy);color:#fff;font-size:12px'>Overall</th>" +
               (($checkNames | ForEach-Object { "<th style='padding:8px 6px;background:$($Brand.Navy);color:#fff;font-size:11px;writing-mode:horizontal-tb'>$_</th>" }) -join '')

    $matRows = foreach ($s in $Servers) {
        $cellOverall = "<td style='text-align:center;padding:6px'>$(ConvertTo-Badge $s.Overall)</td>"
        $cells = foreach ($cn in $checkNames) {
            $c = $s.Checks | Where-Object Name -eq $cn | Select-Object -First 1
            $st = if ($c) { $c.Status } else { 'n/a' }
            $dot = switch ($st) {
                'OK'{$Brand.OkFg} 'WARNING'{$Brand.WarnFg} 'CRITICAL'{$Brand.CritFg} default{$Brand.Muted}
            }
            $glyph = switch ($st) { 'OK'{'&#10003;'} 'WARNING'{'!'} 'CRITICAL'{'&#10007;'} 'ERROR'{'&#10007;'} default{'-'} }
            "<td title='$cn : $st' style='text-align:center;padding:6px;color:$dot;font-weight:700'>$glyph</td>"
        }
        "<tr><td style='padding:6px 10px;font-weight:600;color:$($Brand.Navy)'>$($s.Instance)</td>$cellOverall$($cells -join '')</tr>"
    }

    # ---- Per-server detail ----
    $detail = foreach ($s in $Servers) {
        $band = switch ($s.Overall) { 'OK'{$Brand.OkFg} 'WARNING'{$Brand.WarnFg} default{$Brand.CritFg} }
        $body = if (-not $s.Reachable) {
            "<div style='padding:10px;color:$($Brand.CritFg)'>Unreachable: $([System.Web.HttpUtility]::HtmlEncode($s.Error))</div>"
        } else {
            $blocks = foreach ($c in $s.Checks) {
                $tbl = ConvertTo-DetailTable $c.Rows
                "<div style='margin:8px 0'>" +
                "<div style='font-size:13px'><strong style='color:$($Brand.Navy)'>$($c.Name)</strong> &nbsp; $(ConvertTo-Badge $c.Status) &nbsp; <span style='color:$($Brand.Muted);font-size:12px'>$([System.Web.HttpUtility]::HtmlEncode($c.Summary))</span></div>" +
                $tbl + "</div>"
            }
            ($blocks -join '')
        }
        $svc = if ($s.WindowsSvc) { " &nbsp;<span style='font-size:11px;color:$($Brand.Muted)'>[$($s.WindowsSvc)]</span>" } else { '' }
        "<div style='border:1px solid $($Brand.Border);border-radius:8px;margin:14px 0;overflow:hidden'>" +
        "<div style='background:$($Brand.HeadRow);border-left:6px solid $band;padding:10px 14px'>" +
        "<span style='font-size:16px;font-weight:700;color:$($Brand.Navy)'>$($s.Instance)</span> &nbsp; $(ConvertTo-Badge $s.Overall)$svc</div>" +
        "<div style='padding:6px 14px 12px 14px'>$body</div></div>"
    }

    $crit = @($Servers | Where-Object Overall -eq 'CRITICAL').Count
    $warn = @($Servers | Where-Object Overall -eq 'WARNING').Count
    $ok   = @($Servers | Where-Object Overall -eq 'OK').Count

    @"
<!DOCTYPE html><html><head><meta charset='utf-8'></head>
<body style='margin:0;background:#F4F6F9;font-family:Segoe UI,Arial,sans-serif;color:#1a1a1a'>
<div style='max-width:1100px;margin:0 auto;padding:18px'>
  <div style='background:$($Brand.Navy);color:#fff;padding:18px 22px;border-radius:10px'>
    <div style='font-size:22px;font-weight:700'>SQL Server Health Check &mdash; Consolidated Report</div>
    <div style='font-size:13px;opacity:.85;margin-top:4px'>SQL DBA Champs &bull; Generated $($RunTime.ToString('dddd, dd MMM yyyy HH:mm')) &bull; $(@($Servers).Count) instance(s)</div>
  </div>
  <table style='width:100%;border-collapse:separate;border-spacing:12px 0;margin:14px 0'><tr>
    <td style='background:$($Brand.CritBg);color:$($Brand.CritFg);border-radius:8px;padding:12px;text-align:center;width:33%'><div style='font-size:26px;font-weight:800'>$crit</div><div style='font-size:12px'>CRITICAL</div></td>
    <td style='background:$($Brand.WarnBg);color:$($Brand.WarnFg);border-radius:8px;padding:12px;text-align:center;width:33%'><div style='font-size:26px;font-weight:800'>$warn</div><div style='font-size:12px'>WARNING</div></td>
    <td style='background:$($Brand.OkBg);color:$($Brand.OkFg);border-radius:8px;padding:12px;text-align:center;width:33%'><div style='font-size:26px;font-weight:800'>$ok</div><div style='font-size:12px'>HEALTHY</div></td>
  </tr></table>
  <div style='background:#fff;border:1px solid $($Brand.Border);border-radius:8px;padding:12px;overflow-x:auto'>
    <div style='font-size:14px;font-weight:700;color:$($Brand.Navy);margin-bottom:8px'>Estate summary</div>
    <table style='border-collapse:collapse;width:100%'><thead><tr>$matHead</tr></thead><tbody>$($matRows -join '')</tbody></table>
    <div style='font-size:11px;color:$($Brand.Muted);margin-top:8px'>&#10003; OK &nbsp;&nbsp; ! Warning &nbsp;&nbsp; &#10007; Critical/Error</div>
  </div>
  <div style='font-size:16px;font-weight:700;color:$($Brand.Navy);margin:20px 0 4px'>Per-server detail</div>
  $($detail -join '')
  <div style='color:$($Brand.Muted);font-size:11px;margin-top:18px;text-align:center'>Automated by Invoke-SqlHealthCheck.ps1 &bull; sqldbachamps.com</div>
</div></body></html>
"@
}

# =====================================================================================
# 6. Excel export (ImportExcel if present, else CSV fallback)
# =====================================================================================

function Export-ReportExcel {
    param($Servers, [string]$Path)
    $flat = foreach ($s in $Servers) {
        foreach ($c in $s.Checks) {
            [pscustomobject]@{
                Instance = $s.Instance; Overall = $s.Overall
                Check    = $c.Name; Status = $c.Status; Detail = $c.Summary
            }
        }
    }
    if (Get-Module -ListAvailable -Name ImportExcel) {
        Import-Module ImportExcel -ErrorAction SilentlyContinue
        $flat | Export-Excel -Path $Path -WorksheetName 'Summary' -AutoSize -FreezeTopRow -BoldTopRow -ClearSheet
        foreach ($s in $Servers) {
            foreach ($c in ($s.Checks | Where-Object { @($_.Rows).Count })) {
                $sheet = ("{0}_{1}" -f $s.Instance, ($c.Name -replace '[^\w]','')) 
                $sheet = $sheet.Substring(0, [Math]::Min(31,$sheet.Length))
                $c.Rows | Export-Excel -Path $Path -WorksheetName $sheet -AutoSize -BoldTopRow
            }
        }
        return $Path
    }
    # Fallback: single CSV of the summary
    $csv = [System.IO.Path]::ChangeExtension($Path, 'csv')
    $flat | Export-Csv -Path $csv -NoTypeInformation -Encoding UTF8
    return $csv
}

# =====================================================================================
# 7. Main
# =====================================================================================

$runTime = Get-Date
if (-not (Test-Path $OutputPath)) { New-Item -ItemType Directory -Path $OutputPath | Out-Null }

$instances = Get-InstanceList -Path $ConfigPath -Fallback $DefaultInstances
if (-not $instances) { throw "No instances found. Populate $ConfigPath or `$DefaultInstances." }

Write-Host "Health check starting for $($instances.Count) instance(s)..." -ForegroundColor Cyan
$results = foreach ($i in $instances) {
    Write-Host "  -> $i" -ForegroundColor Gray
    Test-ServerHealth -Instance $i
}

$stamp    = $runTime.ToString('yyyyMMdd_HHmm')
$htmlPath = Join-Path $OutputPath "HealthCheck_$stamp.html"
$xlsPath  = Join-Path $OutputPath "HealthCheck_$stamp.xlsx"

$html = New-HtmlReport -Servers $results -RunTime $runTime
$html | Out-File -FilePath $htmlPath -Encoding UTF8
Write-Host "HTML  : $htmlPath" -ForegroundColor Green

$excel = Export-ReportExcel -Servers $results -Path $xlsPath
Write-Host "Excel : $excel" -ForegroundColor Green

# ---- Consolidated email ----
if (-not $NoEmail) {
    $crit = @($results | Where-Object Overall -eq 'CRITICAL').Count
    $warn = @($results | Where-Object Overall -eq 'WARNING').Count
    $tag  = if ($crit) {"[CRITICAL x$crit]"} elseif ($warn) {"[WARNING x$warn]"} else {'[ALL HEALTHY]'}
    $subject = "SQL Health Check $tag - $($results.Count) servers - $($runTime.ToString('yyyy-MM-dd HH:mm'))"

    $mail = @{
        SmtpServer = $SmtpServer; Port = $SmtpPort
        From = $MailFrom; To = $MailTo
        Subject = $subject; Body = $html; BodyAsHtml = $true
        Attachments = $excel
    }
    if ($UseSsl) { $mail.UseSsl = $true }
    try {
        # Send-MailMessage is deprecated by MS but remains the on-prem SMTP-relay standard.
        Send-MailMessage @mail -ErrorAction Stop
        Write-Host "Email sent to $($MailTo -join ', ')" -ForegroundColor Green
    } catch {
        Write-Warning "Email failed: $($_.Exception.Message). Reports are still on disk."
    }
}

Write-Host "Done." -ForegroundColor Cyan
