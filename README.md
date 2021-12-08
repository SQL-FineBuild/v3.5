# SQL-FineBuild v3.5

SQL FineBuild provides 1-click install and best-practice configuration on Windows of SQL Server 2019 through to SQL Server 2005.

**This is a Beta version of SQL FineBuild v3.5.0.**  This Repository should be used to hold any enhancements from v3.4

To prepare the code for use, download the relevant folder for your version of SQL Server, then download the _Build Scripts_ folder and copy it into the _\FineBuild_ folder.  

The files for each release of SQL Server are found in the _SQL2019_, _SQL2017_, etc folders.  The _Build Scripts_ folder is common to all releases and is in a separate folder.

Documentation is found on the [SQL FineBuild Wiki](https://github.com/SQL-FineBuild/Common/wiki).  Issues shoud be logged in the [Issues List](https://github.com/SQL-FineBuild/Common/issues).

If you are new to SQL FineBuild, please see [SQL FineBuild QuickStart](https://github.com/SQL-FineBuild/Common/wiki/SQL-FineBuild-Quickstart).

Changes compared to v3.4.0:

* Various small bug fixes
* Added parameter /ClusWinSuffix:
* Added parameters /ClusDBIPExtra: /ClusASIPExtra: /ClusRSIPExtra:
* Added support for SQL Server 2019
* Added support for Windows Server 2019, Windows Server 2022 and Windows 11
* Greatly improved support for Availability Groups
