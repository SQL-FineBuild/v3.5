-- Copyright FineBuild Team © 2021.  Distributed under Ms-Pl License
--
-- Rebuild the PolyBase proxy authorisations
-- Based on a script from Microsoft Support, amended to be Idempotent and to use parameter values
--
-- This process is needed for two reasons:
-- 1. The CREATE CERTIFICATE processing appears to be partly asynchronous
--    On a fast server the CREATE CERTIFICATE processing may not complete before the subsequent
--    commands, which prevents them from working.  This condition does not trigger an error in the 
--    Microsoft PolyBase install, so the relevant processing is repeated here to ensure the relevant
--    accounts and permissions are created.
-- 2. The value used as the Certificate Credential must be the same on all servers in the failover set.
--    The Microsoft code generates a unique guid value for each server it is run on, but
--    this causes problems when the servers are in a Distributed Availability Group set or
--    an Availability Group set.  The main symptom is that the Polybase DMS service will fail
--    on secondary servers.  The calling process for this script provides a consistent
--    value for the Certificate Credential and passes this as a parameter to this routine.

USE [DWConfiguration]
GO
IF DATABASEPROPERTYEX('DWConfiguration', 'Updateability') = 'READ_WRITE'
BEGIN;
    IF DATABASE_PRINCIPAL_ID('$(strPBSvcAcnt)') IS NULL
    BEGIN;
        CREATE USER [$(strPBSvcAcnt)] FOR LOGIN [$(strPBSvcAcnt)];
    END;

    ALTER ROLE [db_datareader] ADD MEMBER [$(strPBSvcAcnt)];
    ALTER ROLE [db_datawriter] ADD MEMBER [$(strPBSvcAcnt)];
END;

USE [DWQueue]
GO
IF DATABASEPROPERTYEX('DWQueue', 'Updateability') = 'READ_WRITE'
BEGIN;
    IF DATABASE_PRINCIPAL_ID('$(strPBSvcAcnt)') IS NULL
    BEGIN;
        CREATE USER [$(strPBSvcAcnt)] FOR LOGIN [$(strPBSvcAcnt)];
    END;

    ALTER ROLE [db_datareader] ADD MEMBER [$(strPBSvcAcnt)];
    ALTER ROLE [db_datawriter] ADD MEMBER [$(strPBSvcAcnt)];

    GRANT EXEC ON [DWQueue].dbo.[MessageQueueActivate] TO [$(strPBSvcAcnt)];
    GRANT EXEC ON [DWQueue].dbo.[MessageQueueDeleteMessage] TO [$(strPBSvcAcnt)];
    GRANT EXEC ON [DWQueue].dbo.[MessageQueueDequeueAccept] TO [$(strPBSvcAcnt)];
    GRANT EXEC ON [DWQueue].dbo.[MessageQueueEnqueue] TO [$(strPBSvcAcnt)];
    GRANT EXEC ON [DWQueue].dbo.[MessageQueuePeek] TO [$(strPBSvcAcnt)];
    GRANT EXEC ON [DWQueue].dbo.[MessageQueueReceive] TO [$(strPBSvcAcnt)];
    GRANT EXEC ON [DWQueue].dbo.[TransactionStateDelete] TO [$(strPBSvcAcnt)];
    GRANT EXEC ON [DWQueue].dbo.[TransactionStateGetCommittedOperationState] TO [$(strPBSvcAcnt)];
    GRANT EXEC ON [DWQueue].dbo.[TransactionStateGetCurrentOperationState] TO [$(strPBSvcAcnt)];
    GRANT EXEC ON [DWQueue].dbo.[TransactionStateCreate] TO [$(strPBSvcAcnt)];
    GRANT EXEC ON [DWQueue].dbo.[TransactionStateUpdate] TO [$(strPBSvcAcnt)];
    GRANT EXEC ON [DWQueue].dbo.[MessageQueueUpdate] TO [$(strPBSvcAcnt)];
END;

USE [DWDiagnostics]
GO
IF DATABASEPROPERTYEX('DWDiagnostics', 'Updateability') = 'READ_WRITE'
BEGIN;
    IF DATABASE_PRINCIPAL_ID('$(strPBSvcAcnt)') IS NULL
    BEGIN;
        CREATE USER [$(strPBSvcAcnt)] FOR LOGIN [$(strPBSvcAcnt)];
    END;

    ALTER ROLE [db_datareader] ADD MEMBER [$(strPBSvcAcnt)];
    ALTER ROLE [db_datawriter] ADD MEMBER [$(strPBSvcAcnt)];
    ALTER ROLE [db_ddladmin] ADD MEMBER [$(strPBSvcAcnt)];
END;
--***** 
-- The following section serves as a wrapper of sp_pdw_sm_detach to execute it as SYSADMIN. This procedure is created during setup/upgrade.
--*****

USE [DWConfiguration]
GO
IF DATABASEPROPERTYEX('DWConfiguration', 'Updateability') = 'READ_WRITE'
BEGIN;
    IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[sp_pdw_sm_detach]') AND type in (N'P', N'PC'))
    BEGIN;
        PRINT 'DROP PROCEDURE [sp_pdw_sm_detach]';
        DROP PROCEDURE [sp_pdw_sm_detach];
    END;

    PRINT 'CREATE PROCEDURE [sp_pdw_sm_detach]';
    EXECUTE(N'CREATE PROCEDURE [sp_pdw_sm_detach]
    @FileName nvarchar(45)  /* shared memory name */
    AS
    BEGIN;
        SET NOCOUNT ON;
        EXEC [sp_sm_detach] @FileName;
    END;');
    GRANT EXEC ON [DWConfiguration].dbo.[sp_pdw_sm_detach] TO [$(strPBSvcAcnt)];

    --Create a certificate and sign the procedure with a credential unique to the Failover group
    IF EXISTS (SELECT * FROM [sys].[certificates] WHERE name = N'_##PDW_SmDetachSigningCertificate##') 
    BEGIN;
        PRINT 'DROP CERTIFICATE _##PDW_SmDetachSigningCertificate##';
        DROP CERTIFICATE _##PDW_SmDetachSigningCertificate##;
    END;

    PRINT 'CREATE CERTIFICATE _##PDW_SmDetachSigningCertificate##';
    DECLARE @certpasswordDCB nvarchar(max);
    SET @certpasswordDCB = QUOTENAME(N'$(strPBCertPassword)', N'''');
    EXECUTE(N'CREATE CERTIFICATE _##PDW_SmDetachSigningCertificate## ENCRYPTION BY PASSWORD = ' + @certpasswordDCB + N' WITH  SUBJECT = ''For signing sp_pdw_sm_detach SP'';');
    WAITFOR DELAY '00:00:01';
    EXECUTE(N'ADD SIGNATURE to [sp_pdw_sm_detach] BY CERTIFICATE _##PDW_SmDetachSigningCertificate## WITH PASSWORD=' + @certpasswordDCB + N';');
    ALTER CERTIFICATE _##PDW_SmDetachSigningCertificate## REMOVE PRIVATE KEY;
END;
GO

DECLARE @certBinaryBytes varbinary(max);
SET @certBinaryBytes = CERTENCODED(cert_id('_##PDW_SmDetachSigningCertificate##')); 
DECLARE @cmd nvarchar(max)
SET @cmd = N'use master;
  IF EXISTS (SELECT * FROM [sys].[server_principals] WHERE name = N''l_certSignSmDetach'' AND type = N''C'') 
      BEGIN;
      PRINT ''DROP LOGIN l_certSignSmDetach'';
      DROP LOGIN l_certSignSmDetach;
      END;
  IF EXISTS (SELECT * FROM [sys].[certificates] WHERE name = N''_##PDW_SmDetachSigningCertificate##'') 
      BEGIN;
      PRINT ''DROP CERTIFICATE _##PDW_SmDetachSigningCertificate##'';
      DROP CERTIFICATE _##PDW_SmDetachSigningCertificate##;
      END;
  PRINT ''CREATE CERTIFICATE _##PDW_SmDetachSigningCertificate##'';
  CREATE CERTIFICATE [_##PDW_SmDetachSigningCertificate##] FROM BINARY = ' + sys.fn_varbintohexstr(@certBinaryBytes) + N'
  WAITFOR DELAY ''00:00:01'';
  PRINT ''CREATE LOGIN [l_certSignSmDetach]'';
  CREATE LOGIN [l_certSignSmDetach] FROM CERTIFICATE [_##PDW_SmDetachSigningCertificate##];
  ALTER SERVER ROLE sysadmin ADD MEMBER [l_certSignSmDetach];'
EXEC(@cmd);
GO

--***** 
-- The following section serves as a wrapper of sp_polybase_authorize to execute it as SYSADMIN. This procedure is created during setup/upgrade.
--*****
USE [DWConfiguration]
GO
IF DATABASEPROPERTYEX('DWConfiguration', 'Updateability') = 'READ_WRITE'
BEGIN;
    IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[sp_pdw_polybase_authorize]') AND type in (N'P', N'PC'))
    BEGIN;
        PRINT 'DROP PROCEDURE [sp_pdw_polybase_authorize]';
        DROP PROCEDURE [sp_pdw_polybase_authorize];
    END;

    PRINT 'CREATE PROCEDURE [sp_pdw_polybase_authorize]';
    EXECUTE(N'CREATE PROCEDURE [sp_pdw_polybase_authorize]
    @AppName nvarchar(max)
    AS
    BEGIN;
        SET NOCOUNT ON;
        EXEC [sp_polybase_authorize] @AppName;
    END;');
    GRANT EXEC ON [DWConfiguration].dbo.[sp_pdw_polybase_authorize] TO [$(strPBSvcAcnt)];

    --Create a certificate and sign the procedure with a credential unique to the Failover group
    IF EXISTS (SELECT * FROM [sys].[certificates] WHERE name = N'_##PDW_PolyBaseAuthorizeSigningCertificate##') 
    BEGIN;
        PRINT 'DROP CERTIFICATE _##PDW_PolyBaseAuthorizeSigningCertificate##';
        DROP CERTIFICATE _##PDW_PolyBaseAuthorizeSigningCertificate##;
    END;

    PRINT 'CREATE CERTIFICATE _##PDW_PolyBaseAuthorizeSigningCertificate##';
    DECLARE @certpasswordDCB nvarchar(max);
    SET @certpasswordDCB = QUOTENAME(N'$(strPBCertPassword)', N'''');
    EXECUTE(N'CREATE CERTIFICATE _##PDW_PolyBaseAuthorizeSigningCertificate## ENCRYPTION BY PASSWORD = ' + @certpasswordDCB + N' WITH  SUBJECT = ''For signing sp_pdw_polybase_authorize SP'';');
    WAITFOR DELAY '00:00:01';
    EXECUTE(N'ADD SIGNATURE to [sp_pdw_polybase_authorize] BY CERTIFICATE _##PDW_PolyBaseAuthorizeSigningCertificate## WITH PASSWORD=' + @certpasswordDCB + N';');
    ALTER CERTIFICATE _##PDW_PolyBaseAuthorizeSigningCertificate## REMOVE PRIVATE KEY;
END;
GO

DECLARE @certBinaryBytes varbinary(max);
SET @certBinaryBytes = CERTENCODED(cert_id('_##PDW_PolyBaseAuthorizeSigningCertificate##')); 
DECLARE @cmd nvarchar(max)
SET @cmd = N'use master;
  IF EXISTS (SELECT * FROM [sys].[server_principals] WHERE name = N''l_certSignPolyBaseAuthorize'' AND type = N''C'') 
      BEGIN;
      PRINT ''DROP LOGIN l_certSignPolyBaseAuthorize'';
      DROP LOGIN l_certSignPolyBaseAuthorize;
      END;
  IF EXISTS (SELECT * FROM [sys].[certificates] WHERE name = N''_##PDW_PolyBaseAuthorizeSigningCertificate##'') 
      BEGIN;
      PRINT ''DROP CERTIFICATE _##PDW_PolyBaseAuthorizeSigningCertificate##'';
      DROP CERTIFICATE _##PDW_PolyBaseAuthorizeSigningCertificate##;
      END;
  PRINT ''CREATE CERTIFICATE [_##PDW_PolyBaseAuthorizeSigningCertificate##]'';
  CREATE CERTIFICATE [_##PDW_PolyBaseAuthorizeSigningCertificate##] FROM BINARY = ' + sys.fn_varbintohexstr(@certBinaryBytes) + N'
  WAITFOR DELAY ''00:00:01'';
  PRINT ''CREATE LOGIN [l_certSignPolyBaseAuthorize]'';
  CREATE LOGIN [l_certSignPolyBaseAuthorize] FROM CERTIFICATE [_##PDW_PolyBaseAuthorizeSigningCertificate##];
  ALTER SERVER ROLE sysadmin ADD MEMBER [l_certSignPolyBaseAuthorize];'
EXEC(@cmd);
GO
