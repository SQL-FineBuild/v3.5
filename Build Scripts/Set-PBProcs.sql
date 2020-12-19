-- Copyright FineBuild Team © 2020.  Distributed under Ms-Pl License
--
-- Rebuild the PolyBase proxy authorisations
-- Based on a script from Microsoft Support
--
-- This proces is needed because CREATE CERTIFICATE processing appears to be partly asynchronous, 
-- and on a fast server the CREATE CERTIFICATE processing may not complete before the subsequent
-- commands which prevents them from working.  This condition does not trigger an error in the 
-- Microsoft PolyBase install, so the relevant processing is repeated here.

--Create a certificate and sign the procedure with a unique password
IF EXISTS (SELECT * FROM [sys].[certificates] WHERE name = N'_##PDW_SmDetachSigningCertificate##') DROP CERTIFICATE _##PDW_SmDetachSigningCertificate##;
GO

DECLARE @certpasswordDCB nvarchar(max);
SET @certpasswordDCB = QUOTENAME(newid(), N'''');
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

--Create a certificate and sign the procedure with a unique password
IF EXISTS (SELECT * FROM [sys].[certificates] WHERE name = N'_##PDW_PolyBaseAuthorizeSigningCertificate##') DROP CERTIFICATE _##PDW_PolyBaseAuthorizeSigningCertificate##;
GO

DECLARE @certpasswordDCB nvarchar(max);
SET @certpasswordDCB = QUOTENAME(newid(), N'''');
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
