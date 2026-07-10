# Centralized SQL Server Health Check

**SQL DBA Champs** · one script, all your on-prem instances, one HTML/Excel report, one email.

## Why the central-orchestrator pattern (not per-server Agent jobs)

A per-instance Agent job can't tell you the service is **down** (a dead instance can't run its own job), and stitching many jobs into *one* consolidated email is painful. Instead, one PowerShell script runs from a central monitor box, connects out to every instance, and produces a single rolled-up report. Schedule that one script with a SQL Agent **CmdExec** job or Task Scheduler.

```
  [ Monitor box ]  Invoke-SqlHealthCheck.ps1
        |  reads servers.txt
        |  connects out (ADO.NET, Windows/SQL auth)
        v
  VMSQL01  VMSQL02  VMSQL03 ...   -> collect checks
        |
        v
  HealthCheck_YYYYMMDD_HHMM.html  +  .xlsx   -> ONE email to the DBA team
```

## What it checks (per instance)

| Check | Source | Signals |
|---|---|---|
| SQL Server service status | connection test (+ optional CIM) | reachable = running |
| SQL Server Agent status | `sys.dm_server_services` | Agent Running/Stopped; alerts only if stopped while set to auto-start |
| Disk space | `sys.dm_os_volume_stats` | % free per volume |
| Database backup status | `msdb.dbo.backupset` | stale/missing full & log backups |
| Failed SQL Agent jobs | `msdb` job history | failures in last N hours |
| Blocking sessions | `sys.dm_exec_requests` | blocked SPIDs + wait time |
| Long-running sessions | `sys.dm_exec_requests` | user **queries** actively running over threshold |
| Long-running transactions | `sys.dm_tran_active_transactions` | **open transactions** over threshold, even if the session is idle/sleeping |
| Database status | `sys.databases` | anything not `ONLINE` |
| Error log | `xp_readerrorlog` | error entries in last N hours |
| CPU usage | scheduler ring buffer | SQL process CPU % |
| Memory usage | `sys.dm_os_sys_memory` | OS available MB + memory state |

Every check is graded **OK / WARNING / CRITICAL**; each server rolls up to its worst check.

## Files

- `Invoke-SqlHealthCheck.ps1` — the orchestrator (no modules required; pure ADO.NET)
- `servers.txt` — your instance list, one per line
- `Deploy-HealthCheckAgentJob.sql` — creates the scheduled SQL Agent job
- `Reports/` — generated HTML/Excel land here

## Requirements

- **Windows PowerShell 5.1** (or PowerShell 7). No SqlServer/dbatools module needed.
- Optional: `ImportExcel` module for a real multi-sheet `.xlsx` (`Install-Module ImportExcel`). Without it you get a summary `.csv`.
- On each **target** instance, the running account needs `VIEW SERVER STATE` and read access to `msdb`:
  ```sql
  -- run on every target instance
  CREATE LOGIN [DOMAIN\SqlHealthCheckSvc] FROM WINDOWS;   -- or a SQL login
  GRANT VIEW SERVER STATE TO [DOMAIN\SqlHealthCheckSvc];
  USE msdb; CREATE USER [DOMAIN\SqlHealthCheckSvc] FOR LOGIN [DOMAIN\SqlHealthCheckSvc];
  EXEC sp_addrolemember N'SQLAgentReaderRole', N'DOMAIN\SqlHealthCheckSvc';
  ```
- SMTP relay reachable from the monitor box.

## Run it manually

```powershell
# Generate reports only, no email (great for first run)
.\Invoke-SqlHealthCheck.ps1 -NoEmail

# Deliver through your existing Database Mail profile (recommended - no passwords here)
.\Invoke-SqlHealthCheck.ps1 `
    -UseDatabaseMail -DbMailProfile 'SQLHealthCheckProfile' -DbMailInstance 'localhost' `
    -MailTo 'praveensqldba12@gmail.com'

# Also attach the Excel workbook (watch Database Mail's MaxFileSize, default 1 MB)
.\Invoke-SqlHealthCheck.ps1 -UseDatabaseMail -DbMailProfile 'SQLHealthCheckProfile' -AttachExcel `
    -MailTo 'praveensqldba12@gmail.com'

# SQL authentication to the targets
.\Invoke-SqlHealthCheck.ps1 -SqlAuth -SqlUser hc_reader -SqlPassword '***' -NoEmail
```

### Email delivery

**Database Mail (recommended).** If you've configured a Database Mail profile, use `-UseDatabaseMail -DbMailProfile '<name>'`. The report is sent with `msdb.dbo.sp_send_dbmail` using that profile, so your SMTP host and Gmail app password stay inside SQL's encrypted config - nothing sensitive touches the command line or the Agent job step. `sp_send_dbmail` queues the mail, so a clean run means "accepted"; confirm actual delivery with:
```sql
SELECT TOP 5 sent_status, sent_date, subject FROM msdb.dbo.sysmail_allitems ORDER BY mailitem_id DESC;
```

**Direct SMTP (fallback).** Without Database Mail you can send straight to Gmail/O365, but they require authentication + TLS: pass `-SmtpServer smtp.gmail.com -SmtpPort 587 -SmtpUser you@gmail.com -SmtpPassword '<APP PASSWORD>'`. Gmail needs a 16-char **App Password** (2-Step Verification must be on) - your normal password will be rejected. For scheduled runs, prefer a secure cred file: `Get-Credential | Export-Clixml smtp.cred` then `-SmtpCredentialPath smtp.cred` (readable only by the account that created it).

## Schedule it

Edit the paths/SMTP at the top of `Deploy-HealthCheckAgentJob.sql`, then run it on the **monitor** instance. It creates a daily 07:00 CmdExec job that calls the script. (CmdExec is used instead of the legacy SQLPS PowerShell subsystem, which is a constrained host that breaks modern scripts.)

## Tuning

All thresholds live in the `$Thresholds` hashtable at the top of the script — disk %, backup age, long-run seconds, CPU %, memory MB, lookback windows, connect/query timeouts. Change once, applies to every server. The report palette is in `$Brand`.

## Notes

- Delivery defaults to Database Mail when `-UseDatabaseMail` is set; otherwise it uses `Send-MailMessage` (deprecated by MS but fine for a relay). The script forces TLS 1.2 and auto-enables SSL on any port other than 25, which is what Gmail/O365 require.
- **SQL Agent check** uses `sys.dm_server_services`, so it sees the Agent service state even when the database engine is up. A stopped Agent that's set to **Automatic** startup raises WARNING (flip `$Thresholds.AgentStoppedIsCritical` to `$true` for CRITICAL); a stopped Agent on Manual/Disabled startup is treated as intentional and stays OK. When Agent is down, the Failed Agent Jobs check reports "Agent stopped" instead of a misleading "no failed jobs" (nothing runs to fail).
- **Long-running sessions vs transactions** are two separate checks on purpose. `dm_exec_requests` only shows sessions *actively executing a query right now*, so a `BEGIN TRAN` that never committed (session sleeping / awaiting command) is invisible to it - exactly what `DBCC OPENTRAN` finds. The transactions check keys off `transaction_begin_time` in the tran DMVs, so it catches those idle-but-open transactions (the ones that hold locks and pin the log). Thresholds: `$Thresholds.LongTxnSeconds` / `LongTxnCritSeconds`.
- The disk check reports only volumes that host database files (the ones that matter). Add a CIM/WMI branch if you want every volume.
- **Error log**: searches multiple patterns (login failures, `Error:`, Severity 16-25, `failed`, `corrupt`, `deadlock`, `cannot`) so messages like *"Login failed for user 'DevUser'..."* are captured, and shows the full message text. Severe entries (corruption, Severity 20+, I/O errors 823-825, stack dumps) escalate the check to CRITICAL. Tune `$ErrorLogSearchTerms` / `$ErrorLogCriticalTerms` near the top of the script. Note: `xp_readerrorlog` needs **securityadmin** membership (or `GRANT EXECUTE`) on each target - grant it to the monitoring account alongside `VIEW SERVER STATE`.

---
PMSQLDBA · sqldbachamps.com
