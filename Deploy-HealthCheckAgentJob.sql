/*====================================================================================
  SQL DBA Champs | Deploy the Centralized Health Check as a SQL Agent Job
------------------------------------------------------------------------------------
  Run this ON THE CENTRAL MONITOR INSTANCE (the box that reaches all targets and
  hosts your Database Mail profile).

  It creates a daily job that runs Invoke-SqlHealthCheck.ps1 via a CmdExec step.
  Delivery uses your EXISTING Database Mail profile through sp_send_dbmail - so no
  SMTP server, username, or app password ever appears on the command line or in msdb.

  PREREQUISITES
   1. Copy the toolkit to a local path (see @ScriptDir below).
   2. You already created a Database Mail profile (Management > Database Mail).
      Put its name in @DbMailProfile.
   3. The SQL Agent service account (or CmdExec proxy) needs:
        - Read/execute on the toolkit folder
        - Login + VIEW SERVER STATE + msdb read on every TARGET instance
        - securityadmin (or GRANT EXEC on xp_readerrorlog) on targets, for the error-log check
        - Membership in msdb role DatabaseMailUserRole on THIS instance (to send mail)
   4. Edit the four @variables below.

  ---- PRE-FLIGHT: prove the profile works before scheduling anything ----
    EXEC msdb.dbo.sp_send_dbmail
         @profile_name = N'<YourProfile>',
         @recipients   = N'praveensqldba12@gmail.com',
         @subject      = N'Database Mail test',
         @body         = N'If you can read this, the profile works.';
    -- then check status:
    SELECT TOP 5 sent_status, sent_date, subject
    FROM msdb.dbo.sysmail_allitems ORDER BY mailitem_id DESC;
    -- sent_status should read 'sent'. If 'failed', see:
    SELECT TOP 20 * FROM msdb.dbo.sysmail_event_log ORDER BY log_id DESC;
====================================================================================*/

USE [msdb];
GO

DECLARE @JobName       SYSNAME       = N'DBA - SQL Server Health Check (Consolidated)';
DECLARE @ScriptDir     NVARCHAR(400) = N'E:\1_All_SWs\PS_MonitoringScripts';
DECLARE @DbMailProfile NVARCHAR(200) = N'DBAlerts';         -- <-- your profile name
DECLARE @MailTo        NVARCHAR(400) = N'praveensqldba12@gmail.com';     -- comma-separate for multiple
DECLARE @OwnerLogin    SYSNAME       = SUSER_SNAME();

/* PowerShell command line. Delivery goes through Database Mail on the local instance,
   so we pass -UseDatabaseMail + the profile; no SMTP creds needed. Recipients are
   passed as a comma list to -MailTo (the script joins them for sp_send_dbmail).
   Add -AttachExcel to also attach the .xlsx (mind Database Mail's MaxFileSize, default 1MB). */
DECLARE @Command NVARCHAR(MAX) =
    N'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' + @ScriptDir + N'\Invoke-SqlHealthCheck.ps1"'
  + N' -ConfigPath "'      + @ScriptDir + N'\servers.txt"'
  + N' -OutputPath "'      + @ScriptDir + N'\Reports"'
  + N' -UseDatabaseMail'
  + N' -DbMailProfile "'   + @DbMailProfile + N'"'
  + N' -DbMailInstance "'  + @@SERVERNAME + N'"'
  + N' -MailTo "'          + @MailTo + N'"';

-------------------------------------------------------------------------------------
IF EXISTS (SELECT 1 FROM msdb.dbo.sysjobs WHERE name = @JobName)
    EXEC msdb.dbo.sp_delete_job @job_name = @JobName, @delete_unused_schedule = 1;

EXEC msdb.dbo.sp_add_job
     @job_name = @JobName,
     @enabled  = 1,
     @owner_login_name = @OwnerLogin,
     @description = N'Runs the consolidated SQL Server health check across all instances in servers.txt and emails one report via Database Mail. sqldbachamps.com';

EXEC msdb.dbo.sp_add_jobstep
     @job_name   = @JobName,
     @step_name  = N'Run PowerShell health check',
     @subsystem  = N'CMDEXEC',        -- CmdExec, NOT the SQLPS PowerShell subsystem
     @command    = @Command,
     @on_success_action = 1,
     @on_fail_action    = 2,
     @retry_attempts    = 1,
     @retry_interval    = 2;

/* Schedule: every day at 07:00. Adjust @freq_* as needed. */
EXEC msdb.dbo.sp_add_jobschedule
     @job_name = @JobName,
     @name     = N'Daily 07:00',
     @freq_type = 4,                  -- daily
     @freq_interval = 1,
     @active_start_time = 070000;     -- HHMMSS

EXEC msdb.dbo.sp_add_jobserver @job_name = @JobName, @server_name = N'(LOCAL)';

PRINT 'Created/updated job: ' + @JobName;
PRINT 'Command line:';
PRINT @Command;
GO
