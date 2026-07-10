/*====================================================================================
  SQL DBA Champs | Deploy the Centralized Health Check as a SQL Agent Job
------------------------------------------------------------------------------------
  Run this ON THE CENTRAL MONITOR INSTANCE (the box that reaches all targets).

  It creates a job that calls powershell.exe (CmdExec step) to run
  Invoke-SqlHealthCheck.ps1 once per morning. A CmdExec step is used instead of
  the legacy PowerShell subsystem (SQLPS) because SQLPS is a constrained host and
  often breaks modern scripts.

  PREREQUISITES
   1. Copy the toolkit to a local path, e.g. C:\DBA\SqlHealthCheck\
   2. The SQL Agent service account (or a CmdExec proxy) needs:
        - Read/execute on that folder
        - Network access + login (VIEW SERVER STATE, msdb read) on every target
        - Rights to relay through your SMTP server
   3. Edit @ScriptDir / SMTP / mail params below.
====================================================================================*/

USE [msdb];
GO

DECLARE @JobName    SYSNAME       = N'DBA - SQL Server Health Check (Consolidated)';
DECLARE @ScriptDir  NVARCHAR(400) = N'C:\DBA\SqlHealthCheck';
DECLARE @SmtpServer NVARCHAR(200) = N'smtp.yourcompany.local';
DECLARE @MailFrom   NVARCHAR(200) = N'sql-healthcheck@yourcompany.local';
DECLARE @MailTo     NVARCHAR(400) = N'dba-team@yourcompany.local';   -- comma-separate for multiple
DECLARE @OwnerLogin SYSNAME       = SUSER_SNAME();  -- job owner (defaults to you)

/* Build the powershell command line. -ExecutionPolicy Bypass avoids policy blocks;
   quoting handles paths with spaces. Multiple recipients: pass a comma list. */
DECLARE @Command NVARCHAR(MAX) =
    N'powershell.exe -NoProfile -ExecutionPolicy Bypass -File "' + @ScriptDir + N'\Invoke-SqlHealthCheck.ps1"'
  + N' -ConfigPath "'  + @ScriptDir + N'\servers.txt"'
  + N' -OutputPath "'  + @ScriptDir + N'\Reports"'
  + N' -SmtpServer "'  + @SmtpServer + N'"'
  + N' -MailFrom "'    + @MailFrom   + N'"'
  + N' -MailTo "'      + @MailTo      + N'"';

-------------------------------------------------------------------------------------
IF EXISTS (SELECT 1 FROM msdb.dbo.sysjobs WHERE name = @JobName)
    EXEC msdb.dbo.sp_delete_job @job_name = @JobName, @delete_unused_schedule = 1;

EXEC msdb.dbo.sp_add_job
     @job_name = @JobName,
     @enabled  = 1,
     @owner_login_name = @OwnerLogin,
     @description = N'Runs the consolidated SQL Server health check across all instances in servers.txt and emails one report. sqldbachamps.com';

EXEC msdb.dbo.sp_add_jobstep
     @job_name   = @JobName,
     @step_name  = N'Run PowerShell health check',
     @subsystem  = N'CMDEXEC',        -- CmdExec, NOT the SQLPS PowerShell subsystem
     @command    = @Command,
     @on_success_action = 1,          -- quit reporting success
     @on_fail_action    = 2,          -- quit reporting failure
     @retry_attempts    = 1,
     @retry_interval    = 2;

/* Schedule: every day at 07:00. Adjust @freq_* as needed. */
EXEC msdb.dbo.sp_add_jobschedule
     @job_name = @JobName,
     @name     = N'Daily 07:00',
     @freq_type = 4,                  -- daily
     @freq_interval = 1,
     @active_start_time = 070000;     -- HHMMSS

/* Target the local server as the job host. */
EXEC msdb.dbo.sp_add_jobserver @job_name = @JobName, @server_name = N'(LOCAL)';

PRINT 'Created/updated job: ' + @JobName;
PRINT 'Command line:';
PRINT @Command;
GO
