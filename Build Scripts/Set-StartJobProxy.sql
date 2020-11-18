-- Copyright FineBuild Team © 2016 - 2020.  Distributed under Ms-Pl License
-- Process to create a secure Proxy to allow a user to start a SQL Agent job
--
-- This process can be adapted to allow a non-privilged user to run any privileged action 
--
USE msdb
GO
IF NOT EXISTS (SELECT 1 FROM sys.certificates WHERE name='StartJobProxy')
BEGIN;
  CREATE CERTIFICATE StartJobProxy 
    ENCRYPTION BY PASSWORD='$(strStartJobPassword)'
    WITH
     SUBJECT='StartJobProxy'
    ,START_DATE='2000/01/01'
    ,EXPIRY_DATE='2999/12/31';

  WAITFOR DELAY "00:00:01"

  BACKUP CERTIFICATE [StartJobProxy] 
   TO FILE='$(strDirSystemDataShared)StartJobProxy.cer'
   WITH PRIVATE KEY (
     FILE='$(strDirSystemDataShared)StartJobProxy.pvk',
     ENCRYPTION BY PASSWORD='$(strStartJobPassword)',
     DECRYPTION BY PASSWORD='$(strStartJobPassword)');
END;

IF NOT EXISTS (SELECT 1 FROM sys.sysusers WHERE name='StartJobProxy')
BEGIN;
  CREATE USER [StartJobProxy] 
   FROM CERTIFICATE [StartJobProxy];
END;

GRANT AUTHENTICATE TO [StartJobProxy];
GRANT EXECUTE ON msdb.dbo.sp_start_job TO [StartJobProxy];

/*
USE [UserDB]
GO
CREATE CERTIFICATE [StartJobProxy] 
 FROM FILE='$(strDirSystemDataShared)StartJobProxy.cer'
 WITH PRIVATE KEY (
   FILE='$(strDirSystemDataShared)StartJobProxy.pvk',
   ENCRYPTION BY PASSWORD='$(strStartJobPassword)',
   DECRYPTION BY PASSWORD='$(strStartJobPassword)');
GO
*/

/*
-- Example procedure that uses the StartJobProxy
-- [Job Owner] must be the same account that owns 'User Job'
USE [UserDB]
GO
CREATE PROCEDURE dbo.[spUserDBStartJobProc]
WITH EXECUTE AS [Job Owner]
AS
BEGIN;
  EXEC msdb.dbo.sp_start_job @job_name='User Job';
END;
GO
ADD SIGNATURE TO dbo.[spUserDBStartJobProc]
  BY CERTIFICATE [StartJobProxy]
  WITH PASSWORD='$(strStartJobPassword)';
GO
GRANT EXECUTE ON dbo.[spUserDBStartJobProc] TO [Required User or Role];
EXEC dbo.spUserDBStartJobProc;
*/