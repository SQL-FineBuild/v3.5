-- Copyright FineBuild Team © 2020.  Distributed under Ms-Pl License
-- Code adapted from Microsoft standard jobs and amended to additionally cope with Distributed Availability Groups

USE [msdb]
GO
IF EXISTS(SELECT 1 FROM msdb.dbo.sysjobs WHERE name = 'SSIS Server Maintenance Job')
    EXEC msdb.dbo.sp_delete_job @job_name = 'SSIS Server Maintenance Job'
GO

BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0

IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'SSIS Server Maintenance Job', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'Runs every day. The job removes operation records from the database that are outside the retention window and maintains a maximum number of versions per project.', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'##MS_SSISServerCleanupJobLogin##', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'SSIS Server Operation Records Maintenance', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=3, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=3, 
		@retry_interval=3, 
		@os_run_priority=0, @subsystem=N'TSQL',
		@command=N'DECLARE @role int
    SET @role = (
      SELECT 
       MAX(arps.role) 
      FROM [sys].[availability_groups] ag  
      JOIN [sys].[availability_replicas] arp ON arp.group_id = ag.group_id
      LEFT JOIN [sys].[dm_hadr_availability_replica_states] arps ON arps.group_id = arp.group_id AND arps.replica_id = arp.replica_id
      LEFT JOIN [sys].[availability_groups] agpl ON agpl.name = CASE WHEN ag.is_distributed = 1 THEN arp.replica_server_name ELSE ag.name END
      LEFT JOIN [sys].[dm_hadr_availability_replica_states] arpsl ON arpsl.group_id = agpl.group_id AND arpsl.is_local = 1
      LEFT JOIN [sys].[availability_replicas] arpl ON arpl.group_id = arpsl.group_id AND arpl.replica_id = arpsl.replica_id
      INNER JOIN [sys].[availability_databases_cluster] adc ON adc.[group_id] = arpl.[group_id]
      WHERE adc.database_name = ''SSISDB'')
    IF DB_ID(''SSISDB'') IS NOT NULL AND (@role IS NULL OR @role = 1)
    BEGIN
        EXEC [SSISDB].[internal].[cleanup_server_retention_window]
    END', 
		@database_name=N'msdb', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'SSIS Server Max Version Per Project Maintenance', 
		@step_id=2, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=3, 
		@retry_interval=3, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'DECLARE @role int
    SET @role = (
      SELECT 
       MAX(arps.role) 
      FROM [sys].[availability_groups] ag  
      JOIN [sys].[availability_replicas] arp ON arp.group_id = ag.group_id
      LEFT JOIN [sys].[dm_hadr_availability_replica_states] arps ON arps.group_id = arp.group_id AND arps.replica_id = arp.replica_id
      LEFT JOIN [sys].[availability_groups] agpl ON agpl.name = CASE WHEN ag.is_distributed = 1 THEN arp.replica_server_name ELSE ag.name END
      LEFT JOIN [sys].[dm_hadr_availability_replica_states] arpsl ON arpsl.group_id = agpl.group_id AND arpsl.is_local = 1
      LEFT JOIN [sys].[availability_replicas] arpl ON arpl.group_id = arpsl.group_id AND arpl.replica_id = arpsl.replica_id
      INNER JOIN [sys].[availability_databases_cluster] adc ON adc.[group_id] = arpl.[group_id]
      WHERE adc.database_name = ''SSISDB'')
    IF DB_ID(''SSISDB'') IS NOT NULL AND (@role IS NULL OR @role = 1)
    BEGIN
        EXEC [SSISDB].[internal].[cleanup_server_project_version]
    END', 
		@database_name=N'msdb', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'SSISDB Scheduler', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=1, 
		@freq_subday_interval=0, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20001231, 
		@active_end_date=99991231, 
		@active_start_time=0, 
		@active_end_time=120000
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:
GO

IF EXISTS(SELECT 1 FROM msdb.dbo.sysjobs WHERE name = 'SSIS Failover Monitor Job')
    EXEC msdb.dbo.sp_delete_job @job_name = 'SSIS Failover Monitor Job'
GO

BEGIN TRANSACTION
DECLARE @ReturnCode INT
SELECT @ReturnCode = 0

IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'[Uncategorized (Local)]' AND category_class=1)
BEGIN
EXEC @ReturnCode = msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'[Uncategorized (Local)]'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

END

DECLARE @jobId BINARY(16)
EXEC @ReturnCode =  msdb.dbo.sp_add_job @job_name=N'SSIS Failover Monitor Job', 
		@enabled=1, 
		@notify_level_eventlog=2, 
		@notify_level_email=0, 
		@notify_level_netsend=0, 
		@notify_level_page=0, 
		@delete_level=0, 
		@description=N'Runs every 2 minutes. This job execute master.dbo.sp_ssis_startup if detect AlwaysOn failover on SSISDB.', 
		@category_name=N'[Uncategorized (Local)]', 
		@owner_login_name=N'##MS_SSISServerCleanupJobLogin##', @job_id = @jobId OUTPUT
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback

EXEC @ReturnCode = msdb.dbo.sp_add_jobstep @job_id=@jobId, @step_name=N'AlwaysOn Failover Monitor', 
		@step_id=1, 
		@cmdexec_success_code=0, 
		@on_success_action=1, 
		@on_success_step_id=0, 
		@on_fail_action=2, 
		@on_fail_step_id=0, 
		@retry_attempts=3, 
		@retry_interval=3, 
		@os_run_priority=0, @subsystem=N'TSQL', 
		@command=N'DECLARE @role int
    DECLARE @status tinyint
    SET @role = (
      SELECT 
       MAX(arps.role) 
      FROM [sys].[availability_groups] ag  
      JOIN [sys].[availability_replicas] arp ON arp.group_id = ag.group_id
      LEFT JOIN [sys].[dm_hadr_availability_replica_states] arps ON arps.group_id = arp.group_id AND arps.replica_id = arp.replica_id
      LEFT JOIN [sys].[availability_groups] agpl ON agpl.name = CASE WHEN ag.is_distributed = 1 THEN arp.replica_server_name ELSE ag.name END
      LEFT JOIN [sys].[dm_hadr_availability_replica_states] arpsl ON arpsl.group_id = agpl.group_id AND arpsl.is_local = 1
      LEFT JOIN [sys].[availability_replicas] arpl ON arpl.group_id = arpsl.group_id AND arpl.replica_id = arpsl.replica_id
      INNER JOIN [sys].[availability_databases_cluster] adc ON adc.[group_id] = arpl.[group_id]
      WHERE adc.database_name = ''SSISDB'')
    IF DB_ID(''SSISDB'') IS NOT NULL AND (@role IS NULL OR @role = 1)
    BEGIN
        EXEC [SSISDB].[internal].[refresh_replica_status] @server_name = @@SERVERNAME, @status = @status OUTPUT
        IF @status = 1
            EXEC [SSISDB].[catalog].[startup]
    END', 
		@database_name=N'msdb', 
		@flags=0
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_update_job @job_id = @jobId, @start_step_id = 1
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobschedule @job_id=@jobId, @name=N'Monitor Scheduler', 
		@enabled=1, 
		@freq_type=4, 
		@freq_interval=1, 
		@freq_subday_type=4, 
		@freq_subday_interval=2, 
		@freq_relative_interval=0, 
		@freq_recurrence_factor=0, 
		@active_start_date=20001231, 
		@active_end_date=99991231, 
		@active_start_time=0, 
		@active_end_time=235959
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
EXEC @ReturnCode = msdb.dbo.sp_add_jobserver @job_id = @jobId, @server_name = N'(local)'
IF (@@ERROR <> 0 OR @ReturnCode <> 0) GOTO QuitWithRollback
COMMIT TRANSACTION
GOTO EndSave
QuitWithRollback:
    IF (@@TRANCOUNT > 0) ROLLBACK TRANSACTION
EndSave:
GO