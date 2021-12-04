--  Copyright FineBuild Team © 2018-2021.  Distributed under Ms-Pl License
USE master
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

-- Process FB_GetAGServers Function
IF EXISTS (SELECT 1 FROM INFORMATION_SCHEMA.ROUTINES WHERE ROUTINE_TYPE = 'FUNCTION' AND ROUTINE_NAME = N'FB_GetAGServers')
  DROP FUNCTION dbo.FB_GetAGServers;
GO

CREATE FUNCTION [dbo].[FB_GetAGServers] 
   (@AGName         NVARCHAR(128)
   ,@TargetServer   NVARCHAR(128))

RETURNS @AGServers  TABLE
  (AGName           NVARCHAR(128)
  ,AG_Id            UNIQUEIDENTIFIER
  ,AGType           CHAR(1)
  ,AvailabilityMode INT
  ,RequiredCommit   INT
  ,PrimaryServer    NVARCHAR(128)
  ,PrimaryEndpoint  NVARCHAR(128)
  ,Primary_Id       UNIQUEIDENTIFIER
  ,SecondaryServer  NVARCHAR(128)
  ,SecondaryEndpoint NVARCHAR(128)
  ,Secondary_Id     UNIQUEIDENTIFIER
  ,ServerId         INT
  ,TargetServer     CHAR(1))

AS
-- FB_GetAGServers
--
--  Copyright FineBuild Team © 2019 - 2020.  Distributed under Ms-Pl License
--
-- Get list of Servers and their roles in an Availability Group
-- The routine will work out from the AG Name what type of availability group is involved, and process accordingly
--
-- Syntax: SELECT * FROM dbo.FB_GetAGServers('%','')      to automatically select first Secondary as the TargetServer for each Availability Group
--     or: SELECT * FROM dbo.FB_GetAGServers('AGName','') to automatically select first Secondary as the TargetServer for given Availability Group
--     or: SELECT * FROM dbo.FB_GetAGServers('AGName','TargetServer') if a specific target server is required for given Availability Group
--
-- Date        Name  Comment
-- 12/11/2019  EdV   Initial code
-- 28/01/2020  EdV   Initial FineBuild Version
-- 12/03/2020  EdV   Fixed bug for Distributed Availability Groups
--
BEGIN;

  DECLARE @Parameters TABLE
  (AGName           NVARCHAR(128)
  ,TargetServer     NVARCHAR(128));

  INSERT INTO @Parameters (AGName,TargetServer) 
  VALUES(@AGName,@TargetServer);

  INSERT INTO @AGServers (AGName, AG_Id, AGType, AvailabilityMode, RequiredCommit, PrimaryServer, PrimaryEndpoint, Primary_Id, SecondaryServer,  SecondaryEndpoint, Secondary_Id, ServerId)
  SELECT *,ROW_NUMBER() OVER(PARTITION BY AGName ORDER BY SecondaryServer) AS ServerId
  FROM (SELECT
     ag.name AS AGName
    ,ag.Group_id AS AG_Id
    ,CASE WHEN ag.is_distributed = 1 THEN 'D' WHEN ag.basic_features = 1 THEN 'B' WHEN ag.cluster_type_desc = 'none' THEN 'N' ELSE 'C' END AS AGType
    ,ars.availability_mode
    ,ISNULL(ag.required_synchronized_secondaries_to_commit, 0) AS RequiredCommit
    ,CASE WHEN ag.is_distributed = 1 AND arps.replica_id IS NULL THEN arp.replica_server_name 
          WHEN arps.role = 1                                     THEN arp.replica_server_name END AS PrimaryServer
    ,SUBSTRING(arp.endpoint_url, 7, CHARINDEX('.', arp.endpoint_url) - 7) AS PrimaryEndpoint
    ,arpl.replica_id AS Primary_Id
    ,CASE WHEN ag.is_distributed = 1 AND arss.role = 2  THEN ars.replica_server_name 
          WHEN ag.is_distributed <> 1 AND arss.role = 2 THEN ars.replica_server_name END AS SecondaryServer
    ,SUBSTRING(ars.endpoint_url, 7, CHARINDEX('.', ars.endpoint_url) - 7) AS SecondaryEndpoint
    ,arsl.replica_id AS Secondary_Id
    FROM [sys].[availability_groups] ag  
    JOIN [sys].[availability_replicas] arp ON arp.group_id = ag.group_id
    LEFT JOIN [sys].[dm_hadr_availability_replica_states] arps ON arps.group_id = arp.group_id AND arps.replica_id = arp.replica_id
    LEFT JOIN [sys].[availability_groups] agpl ON agpl.name = arp.replica_server_name
    LEFT JOIN [sys].[dm_hadr_availability_replica_states] arpsl ON arpsl.group_id = agpl.group_id AND arpsl.is_local = 1
    LEFT JOIN [sys].[availability_replicas] arpl ON arpl.group_id = arpsl.group_id AND arpl.replica_id = arpsl.replica_id
    JOIN [sys].[availability_replicas] ars ON ars.group_id = ag.group_id
    LEFT JOIN [sys].[dm_hadr_availability_replica_states] arss ON arss.group_id = ars.group_id AND arss.replica_id = ars.replica_id
    LEFT JOIN [sys].[availability_groups] agsl ON agsl.name = ars.replica_server_name
    LEFT JOIN [sys].[dm_hadr_availability_replica_states] arssl ON arssl.group_id = agsl.group_id AND arssl.is_local = 1
    LEFT JOIN [sys].[availability_replicas] arsl ON arsl.group_id = agsl.group_id AND arsl.replica_id = arssl.replica_id
    ) AS ag
  WHERE PrimaryServer IS NOT NULL AND SecondaryServer IS NOT NULL;

  UPDATE @AGServers SET
   TargetServer = 'Y'
  FROM @Parameters p
  WHERE ServerId IN (SELECT MAX(ServerId) FROM @AGServers WHERE (ServerId = 1) OR (p.TargetServer = SecondaryEndpoint) GROUP BY AGName);

  RETURN;

END;
GO
ALTER AUTHORIZATION ON [dbo].[FB_GetAGServers] TO  SCHEMA OWNER;
GO

-- Process FB_AGFailover Procedure
IF EXISTS (SELECT 1 FROM sys.procedures WHERE name = N'FB_AGFailover')
  DROP PROCEDURE dbo.FB_AGFailover;
GO

CREATE PROC [dbo].[FB_AGFailover]
 @AGName            NVARCHAR(120)      = '%' -- Name of AG for Failover
,@TargetServer      NVARCHAR(128)      = ''  -- Name of desired New Primary server
,@Force             CHAR(1)            = 'N' -- Force failover even if Primary and Secondary not synchronised
,@Execute           CHAR(1)            = 'Y' -- Execute commands
,@RemoteCall        VARCHAR(1)         = 'N' -- Internal Use Only
,@Operation         CHAR(1)            = ''  -- Internal Use Only
AS
-- FB_AGFailover
--
--  Copyright FineBuild Team © 2019 - 2021.  Distributed under Ms-Pl License
--
-- This routine performs a failover of an Availability Group
-- The routine will work out from the AG Name what type of availability group is involved, and process accordingly
--
-- Syntax: EXEC FB_AGFailover @AGName='???'
--     or: EXEC FB_AGFailover @AGName='???', @TargetServer='???' if a specific target server is required
--
-- Do not use any of the parameters marked 'Internal Use Only', they are used within the Main Control Process
--
-- This Proc can be run on either the Primary or a Secondary server in the AG
-- The Main Control Process works out which nodes are Primary and Secondary, and performs the relevant commands on them
--
-- Date        Name  Comment
-- 30/05/2019  EdV   Initial code
-- 25/06/2019  EdV   Added @TargetServer and @Execute logic
-- 28/01/2020  EdV   Initial FineBuild version
-- 18/03/2020  EdV   Added @Force logic
-- 06/05/2020  EdV   Improved synchronisation test logic
-- 02/12/2021  EdV   Improved progress message details
--
BEGIN;
  SET NOCOUNT ON;

  DECLARE
   @Server          NVARCHAR(128)  = @@servername
  ,@SQLText         NVARCHAR(2000) = ''
  ,@NonSync         INT            = -1;

  DECLARE @Parameters TABLE
  (AGName           NVARCHAR(128)
  ,TargetServer     NVARCHAR(128)
  ,Force            CHAR(1)
  ,ExecProcess      CHAR(1)
  ,CRLF             CHAR(2)
  ,RemoteCall       CHAR(1)
  ,Operation        CHAR(1));

  INSERT INTO @Parameters (AGName,TargetServer,Force,ExecProcess,CRLF,RemoteCall,Operation)
  VALUES (
   @AGName
  ,@TargetServer
  ,@Force
  ,@Execute
  ,Char(13) + Char(10)
  ,@RemoteCall
  ,@Operation);
  UPDATE @Parameters SET
   AGName           = REPLACE(REPLACE(AGName, '[',''),']','')
  ,TargetServer     = CASE WHEN p.TargetServer <> '' THEN p.TargetServer ELSE CAST(SERVERPROPERTY('ServerName') AS NVARCHAR(128)) END
  FROM @Parameters p;

  DECLARE @AGServers TABLE
  (AGName           NVARCHAR(128)
  ,AG_Id            UNIQUEIDENTIFIER
  ,AGType           CHAR(1)
  ,AvailabilityMode INT
  ,RequiredCommit   INT
  ,PrimaryServer    NVARCHAR(128)
  ,PrimaryEndpoint  NVARCHAR(128)
  ,Primary_Id       UNIQUEIDENTIFIER
  ,SecondaryServer  NVARCHAR(128)
  ,SecondaryEndpoint NVARCHAR(128)
  ,Secondary_Id     UNIQUEIDENTIFIER
  ,ServerId         INT
  ,TargetServer     CHAR(1));
  INSERT INTO @AGServers
  SELECT AGName,AG_Id,AGType,AvailabilityMode,RequiredCommit,PrimaryServer,PrimaryEndpoint,Primary_Id,SecondaryServer,SecondaryEndpoint,Secondary_Id,ServerId,TargetServer
  FROM dbo.FB_GetAGServers((SELECT AGName FROM @Parameters), (SELECT TargetServer FROM @Parameters));

  IF (SELECT ExecProcess FROM @Parameters) <> 'Y' SELECT * FROM @AGServers;

  IF (SELECT RemoteCall FROM @Parameters) <> 'Y' -- Main Control Process
  BEGIN;
 
    SELECT
     @SQLText          = 'Performing Failover of: ' + p.AGName
    ,@SQLText          = @SQLText + p.CRLF + 'Current Primary Server: '   + a.PrimaryServer
    ,@SQLText          = @SQLText + p.CRLF + 'Current Secondary Server: ' + a.SecondaryServer
    ,@SQLText          = @SQLText + p.CRLF + REPLICATE('*', 40)
    FROM @Parameters p
    JOIN @AGServers a ON a.AGName LIKE p.AGName
    WHERE a.TargetServer = 'Y';
    PRINT @SQLText;
    
    SELECT
     @SQLText       = p.CRLF + 'EXECUTE [' + a.PrimaryServer   + '].master.dbo.FB_AGFailover @AGName=''' + p.AGName + ''', @TargetServer=''' + p.TargetServer + ''', @Execute=''' + p.ExecProcess + ''', @RemoteCall=''Y'',@Operation = ''S'''
    FROM @Parameters p
    JOIN @AGServers a ON a.AGName LIKE p.AGName;
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @SQLText       = p.CRLF + 'EXECUTE [' + a.SecondaryServer + '].master.dbo.FB_AGFailover @AGName=''' + p.AGName + ''', @TargetServer=''' + p.TargetServer + ''', @Execute=''' + p.ExecProcess + ''', @RemoteCall=''Y'',@Operation = ''S'''
    FROM @Parameters p
    JOIN @AGServers a ON a.AGName LIKE p.AGName;
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @SQLText       = p.CRLF + 'EXECUTE [' + a.PrimaryServer   + '].master.dbo.FB_AGFailover @AGName=''' + p.AGName + ''', @TargetServer=''' + p.TargetServer + ''', @Execute=''' + p.ExecProcess + ''', @RemoteCall=''Y'',@Operation = ''R'''
    FROM @Parameters p
    JOIN @AGServers a ON a.AGName LIKE p.AGName;
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;
 
    SELECT
     @SQLText       = p.CRLF + 'EXECUTE [' + a.SecondaryServer + '].master.dbo.FB_AGFailover @AGName=''' + p.AGName + ''', @TargetServer=''' + p.TargetServer + ''', @Execute=''' + p.ExecProcess + ''', @RemoteCall=''Y'',@Operation = ''F'''
    FROM @Parameters p
    JOIN @AGServers a ON a.AGName LIKE p.AGName;
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @SQLText       = p.CRLF + 'EXECUTE [' + a.PrimaryServer   + '].master.dbo.FB_AGFailover @AGName=''' + p.AGName + ''', @TargetServer=''' + p.TargetServer + ''', @Execute=''' + p.ExecProcess + ''', @RemoteCall=''Y'',@Operation = ''A'''
    FROM @Parameters p
    JOIN @AGServers a ON a.AGName LIKE p.AGName
    WHERE a.AvailabilityMode = 0 OR a.AGType = 'D';
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @SQLText       = p.CRLF + 'EXECUTE [' + a.SecondaryServer + '].master.dbo.FB_AGFailover @AGName=''' + p.AGName + ''', @TargetServer=''' + p.TargetServer + ''', @Execute=''' + p.ExecProcess + ''', @RemoteCall=''Y'',@Operation = ''A'''
    FROM @Parameters p
    JOIN @AGServers a ON a.AGName LIKE p.AGName
    WHERE a.AvailabilityMode = 0 OR a.AGType = 'D';
    PRINT @SQLText;
    IF (SELECT ExecProcess FROM @Parameters) = 'Y'  EXECUTE sp_executeSQL @SQLText;

    SELECT
     @SQLText       = REPLICATE('*', 40) +
                      p.CRLF + 'SQL Failover of ' + p.AGName + ' to ' + a.SecondaryServer + ' complete'
    FROM @Parameters p
    JOIN @AGServers a ON a.AGName = p.AGName
    WHERE a.TargetServer = 'Y';
    SELECT
     @SQLText       = @SQLText + p.CRLF + REPLICATE('*', 40)
    ,@SQLText       = @SQLText + p.CRLF 
    ,@SQLText       = @SQLText + p.CRLF + 'Update DNS Alias for ' + p.AGName + ' to point to ' + a.SecondaryServer
    ,@SQLText       = @SQLText + p.CRLF 
    ,@SQLText       = @SQLText + p.CRLF + REPLICATE('*', 40)
    FROM @Parameters p
    JOIN @AGServers a ON a.AGName = p.AGName
    WHERE a.TargetServer = 'Y' AND a.AGType = 'D';
    PRINT @SQLText;

  END;

  -- Start of Utility Functions called from the Main Control Process

  IF (SELECT Operation FROM @Parameters) IN ('S','A') -- AG Communication Mode
  BEGIN; 
    SELECT
     @SQLText       = '/* Processing on ' + @Server + ' */ ' +
                      p.CRLF + 'ALTER AVAILABILITY GROUP [' + p.AGName + '] ' +
                      p.CRLF + '  MODIFY AVAILABILITY GROUP ON ' 
    FROM @Parameters p;
    SELECT
     @SQLText       = @SQLText + p.CRLF + '  ''' + CASE WHEN p.Operation = 'S' THEN ag.PrimaryServer ELSE ag.SecondaryServer END + ''' '
    ,@SQLText       = @SQLText + 'WITH (AVAILABILITY_MODE=' + CASE WHEN p.Operation = 'S' THEN 'SYNCHRONOUS' ELSE 'ASYNCHRONOUS' END + '_COMMIT),'
    FROM @Parameters p
    JOIN @AGServers ag ON ag.AGName LIKE p.AGName AND ag.TargetServer = 'Y';
    SELECT @SQLText = LEFT(@SQLText, LEN(@SQLText) - 1) + ';';
    PRINT @SQLText;
    EXECUTE sp_executeSQL @SQLText;
    PRINT '';
  END;

  IF (SELECT Operation FROM @Parameters) = 'R' -- AG Role Change
  BEGIN; 
  SELECT @SQLText   = '/* Processing on ' + @Server + ' */'
  PRINT @SQLText;
    WHILE (@NonSync <> 0)
    BEGIN;
      SELECT
       @SQLText     = CASE WHEN @NonSync > 0 THEN 'Waiting for Primary states to be SYNCHRONIZED'
                           ELSE '' END;
      IF @SQLText <> '' PRINT @SQLText;
      WAITFOR DELAY '00:00:01'
      SELECT
       d.name
      ,drs.synchronization_state_desc
      ,drs.log_send_queue_size
      FROM @Parameters p
      JOIN @AGServers ag ON ag.AGName LIKE p.AGName AND ag.TargetServer = 'Y'
      JOIN [master].[sys].[dm_hadr_database_replica_states] drs ON drs.replica_id = ag.Primary_Id 
      JOIN [master].[sys].[databases] d ON d.database_id = drs.database_id
      WHERE drs.synchronization_state_desc NOT IN ('SYNCHRONIZED','SYNCHRONIZING') OR drs.log_send_queue_size > 0;
      SELECT
       @NonSync     = CASE WHEN p.Force = 'Y' THEN 0 ELSE @@ROWCOUNT END
      FROM @Parameters p;
    END;

    SELECT
     @SQLText       = '/* Processing on ' + @Server + ' */ ' +
                      p.CRLF + 'ALTER AVAILABILITY GROUP [' + p.AGName + '] SET (ROLE=SECONDARY);' 
    FROM @Parameters p;
    PRINT @SQLText;
    EXECUTE sp_executeSQL @SQLText;
    PRINT '';
  END;

  IF (SELECT Operation FROM @Parameters) = 'F' -- AG Failover
  BEGIN; 
    SELECT 'Before Failover status', * FROM @AGServers;
    SELECT @SQLText = '/* Processing on ' + @Server + ' */'
    PRINT @SQLText;
    WHILE (@NonSync <> 0)
    BEGIN;
      SELECT
       @SQLText     = CASE WHEN @NonSync > 0 THEN 'Waiting for Secondary states to be SYNCHRONIZED'
                           ELSE '' END;
      IF @SQLText <> '' PRINT @SQLText;
      WAITFOR DELAY '00:00:01'
      SELECT
       d.name
      ,drs.synchronization_state_desc
      ,drs.redo_queue_size
      FROM @Parameters p
      JOIN @AGServers ag ON ag.AGName LIKE p.AGName AND ag.TargetServer = 'Y'
      JOIN [master].[sys].[dm_hadr_database_replica_states] drs ON drs.replica_id = ag.Secondary_Id
      JOIN [master].[sys].[databases] d ON d.database_id = drs.database_id
      WHERE drs.synchronization_state_desc NOT IN ('SYNCHRONIZED','SYNCHRONIZING') OR drs.redo_queue_size > 0;
      SELECT
       @NonSync     = CASE WHEN p.Force = 'Y' THEN 0 ELSE @@ROWCOUNT END
      FROM @Parameters p;
    END;
    SELECT
     @SQLText       = '/* Processing on ' + @Server + '*/ ' +
                      p.CRLF + CASE WHEN ag.AGType = 'C' THEN 'ALTER AVAILABILITY GROUP [' + p.AGName + '] FAILOVER'
                                    ELSE 'ALTER AVAILABILITY GROUP [' + p.AGName + '] FORCE_FAILOVER_ALLOW_DATA_LOSS' END
    FROM @Parameters p
    JOIN @AGServers ag ON ag.AGName = p.AGName AND ag.TargetServer = 'Y';
    PRINT @SQLText;
    EXECUTE sp_executeSQL @SQLText;
    PRINT '';
  END;

END;

GO
ALTER AUTHORIZATION ON [dbo].[FB_AGFailover] TO  SCHEMA OWNER;
GO

-- Create table for Database Owner Mappings

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FB_AGPostFailoverDBOwner]') AND type in (N'U'))
  DROP TABLE [dbo].[FB_AGPostFailoverDBOwner];
GO

CREATE TABLE [dbo].[FB_AGPostFailoverDBOwner]
([Id]          INTEGER IDENTITY(1,1)
,[DBName]      NVARCHAR(128) NOT NULL
,[DBOwner]     NVARCHAR(128) NOT NULL
,CONSTRAINT [PK_AGPostFailoverDBOwner] PRIMARY KEY CLUSTERED ([Id] ASC));
CREATE UNIQUE NONCLUSTERED INDEX [IX_FB_AGPostFailoverDBOwner] ON [dbo].[FB_AGPostFailoverDBOwner]
([DBNAME] ASC
,[DBOwner] ASC);
GO

-- Create default DB Owner mappings

INSERT INTO [dbo].[FB_AGPostFailoverDBOwner] (DBName, DBOwner) 
  VALUES('master', 'sa');
INSERT INTO [dbo].[FB_AGPostFailoverDBOwner] (DBName, DBOwner) 
  VALUES('model', 'sa');
INSERT INTO [dbo].[FB_AGPostFailoverDBOwner] (DBName, DBOwner) 
  VALUES('msdb', 'sa');
INSERT INTO [dbo].[FB_AGPostFailoverDBOwner] (DBName, DBOwner) 
  VALUES('tempdb', 'sa');
INSERT INTO [dbo].[FB_AGPostFailoverDBOwner] (DBName, DBOwner) 
  VALUES('mssqlsystemresource', 'sa');
INSERT INTO [dbo].[FB_AGPostFailoverDBOwner] (DBName, DBOwner) 
  VALUES('SSISDB', 'sa');
INSERT INTO [dbo].[FB_AGPostFailoverDBOwner] (DBName, DBOwner) 
  SELECT name, 'sa' FROM master.sys.databases WHERE is_distributor = 1;
INSERT INTO [dbo].[FB_AGPostFailoverDBOwner] (DBName, DBOwner) 
  VALUES('DQS_MAIN', '##MS_dqs_db_owner_login##');
INSERT INTO [dbo].[FB_AGPostFailoverDBOwner] (DBName, DBOwner) 
  VALUES('DQS_PROJECTS', '##MS_dqs_db_owner_login##');
INSERT INTO [dbo].[FB_AGPostFailoverDBOwner] (DBName, DBOwner) 
  VALUES('DQS_STAGING_DATA', '##MS_dqs_db_owner_login##');

-- Create table for Database User Mappings

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[FB_AGPostFailoverDBUser]') AND type in (N'U'))
  DROP TABLE [dbo].[FB_AGPostFailoverDBUser];
GO

CREATE TABLE [dbo].[FB_AGPostFailoverDBUser]
([Id]          INTEGER IDENTITY(1,1)
,[DBName]      NVARCHAR(128) NOT NULL
,[DBUser]      NVARCHAR(128) NOT NULL
,[Login]       NVARCHAR(128) NOT NULL
,CONSTRAINT [PK_AGPostFailoverDBUser] PRIMARY KEY CLUSTERED ([Id] ASC));
CREATE UNIQUE NONCLUSTERED INDEX [IX_FB_AGPostFailoverDBUser] ON [dbo].[FB_AGPostFailoverDBUser]
([DBNAME] ASC
,[DBUser] ASC
,[Login] ASC);
GO

-- Create default DB User mappings

INSERT INTO [dbo].[FB_AGPostFailoverDBUser] (DBName, DBUser, Login) 
  VALUES('DQS_MAIN', 'dqs_service', '##MS_dqs_service_login##');
INSERT INTO [dbo].[FB_AGPostFailoverDBUser] (DBName, DBUser, Login) 
  VALUES('DQS_PROJECTS', 'dqs_service', '##MS_dqs_service_login##');
INSERT INTO [dbo].[FB_AGPostFailoverDBUser] (DBName, DBUser, Login) 
  VALUES('DQS_STAGING_DATA', 'dqs_service', '##MS_dqs_service_login##');
INSERT INTO [dbo].[FB_AGPostFailoverDBUser] (DBName, DBUser, Login) 
  VALUES('SSISDB', '##MS_SSISServerCleanupJobUser##', '##MS_SSISServerCleanupJobLogin##');

-- Process FB_AGPostFailover Procedure
IF EXISTS (SELECT 1 FROM sys.procedures WHERE name = N'FB_AGPostFailover')
  DROP PROCEDURE dbo.FB_AGPostFailover;
GO

CREATE PROC [dbo].[FB_AGPostFailover]
AS
-- FB_AGPostFailover
--
--  Copyright FineBuild Team © 2020.  Distributed under Ms-Pl License
--
-- This routine performs post-failover tasks for an Availability Group
--
-- Syntax: EXEC FB_AGPostFailover
--
-- Date        Name  Comment
-- 25/06/2019  EdV   Initial code
-- 28/01/2020  EdV   Initial FineBuild Version
-- 02/04/2020  EdV   Replaced fixed code with lookups to FB_AGPostFailoverDBOwner, FB_AGPostFailoverDBUser
--
BEGIN;
  SET NOCOUNT ON;

  DECLARE 
   @AGName          NVARCHAR(400)
  ,@AlertName       NVARCHAR(400)
  ,@DBName          NVARCHAR(400)
  ,@DBUser          NVARCHAR(400)
  ,@Login           NVARCHAR(400)
  ,@Owner           NVARCHAR(400)
  ,@Role            NVARCHAR(400)
  ,@Enabled         INT
  ,@ScheduleId      VARCHAR(8)
  ,@ServerName      NVARCHAR(400)
  ,@SQLText         NVARCHAR(4000);

  WAITFOR DELAY '00:01'

  SELECT
   @ServerName     = UPPER(CAST(ServerProperty('ServerName') AS NVARCHAR(400)));

  -- Update Log Usage Alerts for Primary/Secondary Server
  DECLARE Log_Alerts CURSOR FAST_FORWARD FOR
  SELECT
   a.name
  ,a.enabled
  ,UPPER(rs.role_desc) AS role_desc
  FROM msdb.dbo.sysalerts a
  CROSS JOIN sys.dm_hadr_availability_replica_states rs
  LEFT JOIN sys.availability_replicas ar
    ON rs.replica_id = ar.replica_id AND rs.group_id = ar.group_id
  WHERE UPPER(ar.replica_server_name) = @ServerName
  AND a.name LIKE 'DB %: Log Usage%'
  AND HAS_DBACCESS(SUBSTRING(LEFT(performance_condition, CHARINDEX('|>|', performance_condition) - 1), LEN('SQLServer:Databases|Percent Log Used|') + 1, LEN(performance_condition))) = 1
  ORDER BY a.name;

  OPEN Log_Alerts;
  FETCH NEXT FROM Log_Alerts INTO @AlertName, @Enabled, @Role;
  PRINT 'Server: ' + @ServerName + ' role: ' + @Role;
  WHILE @@FETCH_STATUS = 0  
  BEGIN;
    SELECT 
     @SQLText = 'EXEC msdb.dbo.sp_update_alert @name=''' + @AlertName + ''''
    ,@SQLText = @SQLText + CASE WHEN @Role = 'PRIMARY'   THEN ',@enabled=1;'
                                WHEN @Role = 'SECONDARY' THEN ',@enabled=0;'
                                ELSE ',@enabled=' + Cast(@Enabled AS NVarchar(2)) + ';' END;
    PRINT @SQLText;
    EXEC sp_executeSQL @SQLText;
    FETCH NEXT FROM Log_Alerts INTO @AlertName, @Enabled, @Role;
  END;
  CLOSE Log_Alerts;
  DEALLOCATE Log_Alerts;

  -- Update Job Schedules for Primary/Secondary Server
  DECLARE Job_Schedules CURSOR FAST_FORWARD FOR
  SELECT
   CAST(s.schedule_id AS varchar(8))
  ,s.enabled
  ,UPPER(rs.role_desc) AS role_desc
  FROM msdb.dbo.sysschedules s
  CROSS JOIN sys.dm_hadr_availability_replica_states rs
  LEFT JOIN sys.availability_replicas ar
    ON rs.replica_id = ar.replica_id AND rs.group_id = ar.group_id
  WHERE UPPER(ar.replica_server_name) = @ServerName
  ORDER BY s.schedule_id;

  OPEN Job_Schedules;
  FETCH NEXT FROM Job_Schedules INTO @ScheduleId, @Enabled, @Role;
  PRINT 'Server: ' + @ServerName + ' role: ' + @Role;
  WHILE @@FETCH_STATUS = 0  
  BEGIN;
    SELECT 
     @SQLText = 'EXEC msdb.dbo.sp_update_schedule @schedule_id=' + @ScheduleId
    ,@SQLText = @SQLText + CASE WHEN @Role = 'PRIMARY'   THEN ',@enabled=1;'
                                WHEN @Role = 'SECONDARY' THEN ',@enabled=0;'
                                ELSE ',@enabled=' + Cast(@Enabled AS NVarchar(2)) + ';' END;
    PRINT @SQLText;
    EXEC sp_executeSQL @SQLText;
    FETCH NEXT FROM Job_Schedules INTO @ScheduleId, @Enabled, @Role;
  END;
  CLOSE Job_Schedules;
  DEALLOCATE Job_Schedules;

  -- Update AG Synchronisation for Primary/Secondary Server
  DECLARE AG_DBNames CURSOR FAST_FORWARD FOR
  SELECT
   ag.name 
  ,db.name AS database_name
  ,UPPER(rs.role_desc) AS role_desc
  ,CASE WHEN do.DBName IS NOT NULL THEN do.DBOwner
        ELSE c.credential_identity END AS DBOwner
  FROM master.sys.availability_groups_cluster ag
  JOIN master.sys.dm_hadr_availability_replica_states rs
    ON rs.group_id = ag.group_id
  JOIN sys.availability_replicas ar
    ON rs.replica_id = ar.replica_id AND rs.group_id = ar.group_id
  JOIN master.sys.availability_databases_cluster dbc
    ON dbc.group_id = ag.group_id
  JOIN master.sys.databases db
    ON db.group_database_id = dbc.group_database_id
  LEFT OUTER JOIN sys.credentials c
    ON c.name = 'StandardDBOwner'
  LEFT OUTER JOIN dbo.FB_AGPostFailoverDBOwner do
    ON do.DBName = db.name
  WHERE UPPER(ar.replica_server_name) = @ServerName
  ORDER BY ag.name, db.name;

  OPEN AG_DBNames;
  FETCH NEXT FROM AG_DBNames INTO @AGName, @DBName, @Role, @Owner;
  WHILE @@FETCH_STATUS = 0  
  BEGIN;
    SELECT @SQLText = CASE WHEN @Role = 'SECONDARY' THEN 'ALTER DATABASE [' + @DBName + '] SET HADR RESUME;'
                           ELSE 'ALTER AUTHORIZATION ON DATABASE::[' + @DBName + '] TO [' + @Owner + '];' END;
    PRINT @SQLText;
    EXEC sp_executeSQL @SQLText;
    FETCH NEXT FROM AG_DBNames INTO @AGName, @DBName, @Role, @Owner;
  END;
  CLOSE AG_DBNames;
  DEALLOCATE AG_DBNames;

  -- Update User Mappings for new Primary Server
  DECLARE AG_DBUsers CURSOR FAST_FORWARD FOR
  SELECT
   DBName 
  ,DBUser
  ,Login 
  FROM master.dbo.FB_AGPostFailoverDBUser
  ORDER BY DBName, DBUser;

  OPEN AG_DBUsers;
  FETCH NEXT FROM AG_DBUsers INTO @DBName, @DBUser, @Login;
  WHILE @@FETCH_STATUS = 0  
  BEGIN;
    IF EXISTS (SELECT 1 FROM master.sys.databases WHERE name = @DBName AND @Role = 'PRIMARY')
    BEGIN;
    SELECT 
     @SQLText       = 'USE [' + @DBName + '];IF EXISTS (SELECT 1 FROM sys.sysusers WHERE name = ''' + @DBUser + ''' AND islogin = 1) '
    ,@SQLText       = @SQLText + 'ALTER USER [' + @DBUser + '] WITH LOGIN = [' + @Login + '];'
    PRINT @SQLText;
    EXEC sp_executeSQL @SQLText;
    END;
    FETCH NEXT FROM AG_DBUsers INTO @DBName, @DBUser, @Login;;
  END;
  CLOSE AG_DBUsers;
  DEALLOCATE AG_DBUsers;

END;
GO

-- Process Job to run FB_AGPostFailover
DECLARE 
  @jobId            BINARY(16);

IF NOT EXISTS (SELECT name FROM msdb.dbo.syscategories WHERE name=N'DBA Tasks' AND category_class=1)
  EXEC msdb.dbo.sp_add_category @class=N'JOB', @type=N'LOCAL', @name=N'DBA Tasks';

IF EXISTS (SELECT 1 FROM msdb.dbo.sysjobs WHERE name = N'DBA: AG State Change')
  EXEC msdb.dbo.sp_delete_job @job_name=N'DBA: AG State Change', @delete_unused_schedule=1;

EXEC msdb.dbo.sp_add_job
 @job_name=N'DBA: AG State Change'
,@enabled=1
,@notify_level_eventlog=0
,@notify_level_email=0
,@notify_level_netsend=0
,@notify_level_page=0
,@delete_level=0
,@description=N'Perform required AG Maintenance when AG state changes'
,@category_name=N'DBA Tasks'
,@owner_login_name=N'sa'
,@job_id = @jobId OUTPUT;

EXEC msdb.dbo.sp_add_jobstep
 @job_id=@jobId
,@step_name=N'AG Post Failover'
,@step_id=1
,@cmdexec_success_code=0
,@on_success_action=1
,@on_success_step_id=0
,@on_fail_action=2
,@on_fail_step_id=0
,@retry_attempts=0
,@retry_interval=0
,@os_run_priority=0
,@subsystem=N'TSQL'
,@command=N'EXEC FB_AGPostFailover'
,@database_name=N'master'
,@flags=4;

EXEC msdb.dbo.sp_update_job
 @job_id = @jobId
,@start_step_id = 1;

EXEC msdb.dbo.sp_add_jobschedule
 @job_id=@jobId
,@name=N'AG State Change - Startup Check'
,@enabled=1
,@freq_type=64
,@freq_interval=0
,@freq_subday_type=0
,@freq_subday_interval=0
,@freq_relative_interval=0
,@freq_recurrence_factor=0
,@active_start_date=20000101
,@active_end_date=99991231
,@active_start_time=0
,@active_end_time=235959;

EXEC msdb.dbo.sp_add_jobserver
 @job_id = @jobId
,@server_name = N'(local)';

WAITFOR DELAY '00:00:02' -- Give SQL Agent time to catch up

-- Process Alerts to trigger 'DBA: AG State Change' Job

IF EXISTS (SELECT 1 FROM msdb.dbo.sysalerts WHERE name=N'Event - AG Failover')
  EXEC msdb.dbo.sp_delete_alert @name=N'Event - AG Failover';

EXEC msdb.dbo.sp_add_alert @name=N'Event - AG Failover', 
		@message_id=1480, 
		@severity=0, 
		@enabled=1, 
		@delay_between_responses=60, 
		@include_event_description_in=0, 
		@category_name=N'[Uncategorized]', 
		@job_name=N'DBA: AG State Change';

EXEC msdb.dbo.sp_add_notification @alert_name=N'Event - AG Failover', @operator_name=N'SQL Alerts', @notification_method=1;

IF EXISTS (SELECT 1 FROM msdb.dbo.sysalerts WHERE name=N'AG State Change')
  EXEC msdb.dbo.sp_delete_alert @name=N'Event - AG State Change';

EXEC msdb.dbo.sp_add_alert @name=N'Event - AG State Change', 
		@message_id=35264, 
		@severity=0, 
		@enabled=1, 
		@delay_between_responses=60, 
		@include_event_description_in=0, 
		@category_name=N'[Uncategorized]', 
		@job_name=N'DBA: AG State Change';

EXEC msdb.dbo.sp_add_notification @alert_name=N'Event - AG State Change', @operator_name=N'SQL Alerts', @notification_method=1;

GO