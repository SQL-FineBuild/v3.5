-- Copyright FineBuild Team © 2021.  Distributed under Ms-Pl License
--
-- Rebuild the PolyBase proxy authorisations
-- Based on a script from Microsoft Support
--
-- This proces is needed because CREATE CERTIFICATE processing appears to be partly asynchronous, 
-- and on a fast server the CREATE CERTIFICATE processing may not complete before the subsequent
-- commands which prevents them from working.  This condition does not trigger an error in the 
-- Microsoft PolyBase install, so the relevant processing is repeated here.

USE [DWConfiguration]
CREATE USER [$(strPBSvcAcnt)] FOR LOGIN [$(strPBSvcAcnt)]
ALTER ROLE [db_datareader] ADD MEMBER [$(strPBSvcAcnt)]
ALTER ROLE [db_datawriter] ADD MEMBER [$(strPBSvcAcnt)]

USE [DWQueue]
CREATE USER [$(strPBSvcAcnt)] FOR LOGIN [$(strPBSvcAcnt)]
ALTER ROLE [db_datareader] ADD MEMBER [$(strPBSvcAcnt)]
ALTER ROLE [db_datawriter] ADD MEMBER [$(strPBSvcAcnt)]

GRANT EXEC ON [DWQueue].dbo.[MessageQueueActivate] TO [$(strPBSvcAcnt)]
GRANT EXEC ON [DWQueue].dbo.[MessageQueueDeleteMessage] TO [$(strPBSvcAcnt)]
GRANT EXEC ON [DWQueue].dbo.[MessageQueueDequeueAccept] TO [$(strPBSvcAcnt)]
GRANT EXEC ON [DWQueue].dbo.[MessageQueueEnqueue] TO [$(strPBSvcAcnt)]
GRANT EXEC ON [DWQueue].dbo.[MessageQueuePeek] TO [$(strPBSvcAcnt)]
GRANT EXEC ON [DWQueue].dbo.[MessageQueueReceive] TO [$(strPBSvcAcnt)]
GRANT EXEC ON [DWQueue].dbo.[TransactionStateDelete] TO [$(strPBSvcAcnt)]
GRANT EXEC ON [DWQueue].dbo.[TransactionStateGetCommittedOperationState] TO [$(strPBSvcAcnt)]
GRANT EXEC ON [DWQueue].dbo.[TransactionStateGetCurrentOperationState] TO [$(strPBSvcAcnt)]
GRANT EXEC ON [DWQueue].dbo.[TransactionStateCreate] TO [$(strPBSvcAcnt)]
GRANT EXEC ON [DWQueue].dbo.[TransactionStateUpdate] TO [$(strPBSvcAcnt)]
GRANT EXEC ON [DWQueue].dbo.[MessageQueueUpdate] TO [$(strPBSvcAcnt)]


USE [DWDiagnostics]
CREATE USER [$(strPBSvcAcnt)] FOR LOGIN [$(strPBSvcAcnt)]
ALTER ROLE [db_datareader] ADD MEMBER [$(strPBSvcAcnt)]
ALTER ROLE [db_datawriter] ADD MEMBER [$(strPBSvcAcnt)]
ALTER ROLE [db_ddladmin] ADD MEMBER [$(strPBSvcAcnt)]

--***** 
-- The following section serves as a wrapper of sp_pdw_sm_detach to execute it as SYSADMIN. This procedure is created during setup/upgrade.
--*****

USE [DWConfiguration]
GO

IF EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[sp_pdw_sm_detach]') AND type in (N'P', N'PC'))
  DROP PROCEDURE [sp_pdw_sm_detach];
GO

CREATE PROCEDURE [sp_pdw_sm_detach]
    -- Parameters
    @FileName nvarchar(45)  -- shared memory name
    AS
        SET NOCOUNT ON;
        EXEC [sp_sm_detach] @FileName;
GO

GRANT EXEC ON [DWConfiguration].dbo.[sp_pdw_sm_detach] TO [$(strPBSvcAcnt)]
GO

--Create a certificate and sign the procedure with a password unique to the Failover group
IF EXISTS (SELECT * FROM [sys].[certificates] WHERE name = N'_##PDW_SmDetachSigningCertificate##') DROP CERTIFICATE _##PDW_SmDetachSigningCertificate##;
GO

DECLARE @certpasswordDCB nvarchar(max);
SET @certpasswordDCB = QUOTENAME($(strPBCertPassword), N'''');
EXECUTE(N'CREATE CERTIFICATE _##PDW_SmDetachSigningCertificate## ENCRYPTION BY PASSWORD = ' + @certpasswordDCB + N' WITH  SUBJECT = ''For signing sp_pdw_sm_detach SP'';');

EXECUTE(N'ADD SIGNATURE to [sp_pdw_sm_detach] BY CERTIFICATE _##PDW_SmDetachSigningCertificate## WITH PASSWORD=' + @certpasswordDCB + N';');
GO
WAITFOR DELAY '00:00:01';
ALTER CERTIFICATE _##PDW_SmDetachSigningCertificate## REMOVE PRIVATE KEY;
GO

DECLARE @certBinaryBytes varbinary(max);
SET @certBinaryBytes = CERTENCODED(cert_id('_##PDW_SmDetachSigningCertificate##')); 
DECLARE @cmd nvarchar(max)
SET @cmd = N'use master;
IF EXISTS (SELECT * FROM [sys].[server_principals] WHERE name = N''l_certSignSmDetach'' AND type = N''C'') DROP LOGIN l_certSignSmDetach;
IF EXISTS (SELECT * FROM [sys].[certificates] WHERE name = N''_##PDW_SmDetachSigningCertificate##'') DROP CERTIFICATE _##PDW_SmDetachSigningCertificate##;
CREATE CERTIFICATE [_##PDW_SmDetachSigningCertificate##] FROM BINARY = ' + sys.fn_varbintohexstr(@certBinaryBytes) + N'
WAITFOR DELAY ''00:00:01'';
CREATE LOGIN [l_certSignSmDetach] FROM CERTIFICATE [_##PDW_SmDetachSigningCertificate##];
ALTER SERVER ROLE sysadmin ADD MEMBER [l_certSignSmDetach];'
EXEC(@cmd)
GO

--***** 
-- The following section serves as a wrapper of sp_polybase_authorize to execute it as SYSADMIN. This procedure is created during setup/upgrade.
--*****
USE [DWConfiguration]
GO

IF  EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[sp_pdw_polybase_authorize]') AND type in (N'P', N'PC'))
    DROP PROCEDURE [sp_pdw_polybase_authorize];
GO

CREATE PROCEDURE [sp_pdw_polybase_authorize]
    -- Parameters
    @AppName nvarchar(max)
    AS
        SET NOCOUNT ON;
        EXEC [sp_polybase_authorize] @AppName;
GO

GRANT EXEC ON [DWConfiguration].dbo.[sp_pdw_polybase_authorize] TO [$(strPBSvcAcnt)];
GO

--Create a certificate and sign the procedure with a password unique to the Failover group
IF EXISTS (SELECT * FROM [sys].[certificates] WHERE name = N'_##PDW_PolyBaseAuthorizeSigningCertificate##') DROP CERTIFICATE _##PDW_PolyBaseAuthorizeSigningCertificate##;
GO

DECLARE @certpasswordDCB nvarchar(max);
SET @certpasswordDCB = QUOTENAME($(strPBCertPassword), N'''');
EXECUTE(N'CREATE CERTIFICATE _##PDW_PolyBaseAuthorizeSigningCertificate## ENCRYPTION BY PASSWORD = ' + @certpasswordDCB + N' WITH  SUBJECT = ''For signing sp_pdw_polybase_authorize SP'';');
EXECUTE(N'ADD SIGNATURE to [sp_pdw_polybase_authorize] BY CERTIFICATE _##PDW_PolyBaseAuthorizeSigningCertificate## WITH PASSWORD=' + @certpasswordDCB + N';');
GO
WAITFOR DELAY '00:00:01';
ALTER CERTIFICATE _##PDW_PolyBaseAuthorizeSigningCertificate## REMOVE PRIVATE KEY;
GO

DECLARE @certBinaryBytes varbinary(max);
SET @certBinaryBytes = CERTENCODED(cert_id('_##PDW_PolyBaseAuthorizeSigningCertificate##')); 
DECLARE @cmd nvarchar(max)
SET @cmd = N'use master;
IF EXISTS (SELECT * FROM [sys].[server_principals] WHERE name = N''l_certSignPolyBaseAuthorize'' AND type = N''C'') DROP LOGIN l_certSignPolyBaseAuthorize;
IF EXISTS (SELECT * FROM [sys].[certificates] WHERE name = N''_##PDW_PolyBaseAuthorizeSigningCertificate##'') DROP CERTIFICATE _##PDW_PolyBaseAuthorizeSigningCertificate##;
CREATE CERTIFICATE [_##PDW_PolyBaseAuthorizeSigningCertificate##] FROM BINARY = ' + sys.fn_varbintohexstr(@certBinaryBytes) + N'
WAITFOR DELAY ''00:00:01'';
CREATE LOGIN [l_certSignPolyBaseAuthorize] FROM CERTIFICATE [_##PDW_PolyBaseAuthorizeSigningCertificate##];
ALTER SERVER ROLE sysadmin ADD MEMBER [l_certSignPolyBaseAuthorize];'
EXEC(@cmd)
GO
