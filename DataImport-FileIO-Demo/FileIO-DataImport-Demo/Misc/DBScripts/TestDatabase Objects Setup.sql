
/**************************************************************************8*/
-- CREATE THE ETL IMPORT TABLE
/**************************************************************************8*/
IF OBJECT_ID('[dbo].[Company_Import]') IS NOT NULL BEGIN
	DROP TABLE [dbo].[Company_Import]
END
GO

SET ANSI_NULLS ON
GO

SET QUOTED_IDENTIFIER ON
GO

SET ANSI_PADDING ON
GO

CREATE TABLE [dbo].[Company_Import](
	[Id] [int] IDENTITY(1,1) NOT NULL,
	[CompanyId] [int] NULL,
	[StartDate] [datetime] NULL,
	[EndDate] [datetime] NULL,
	[LegalName] [varchar](150) NULL,
	[DBAName] [varchar](150) NULL,
	[ChangeDate] [datetime] NULL,
	[UserId] [varchar](30) NULL,
	CONSTRAINT [PK_Company_Import] PRIMARY KEY CLUSTERED (
		[Id] ASC
	)
) 

GO

/**************************************************************************8*/
-- CREATE THE ETL IMPORT TABLE TYPE
/**************************************************************************8*/
IF OBJECT_ID('[dbo].[ImportCompanies]') IS NOT NULL BEGIN 
	DROP PROCEDURE [dbo].[ImportCompanies]
END
GO

IF TYPE_ID('dbo.Company_Import_tt') IS NOT NULL BEGIN
	DROP TYPE dbo.Company_Import_tt
END
GO

CREATE TYPE dbo.Company_Import_tt AS TABLE(
	 [CompanyId] int NOT NULL
	,[StartDate] datetime NULL
	,[EndDate] datetime NULL
	,[LegalName] varchar(150) NULL
	,[DBAName] varchar(150) NULL
	,[ChangeDate] datetime NULL
	,[UserId] varchar(30) NULL
);
GO

IF OBJECT_ID('[dbo].[ImportCompanies]') IS NOT NULL BEGIN 
	DROP PROCEDURE [dbo].[ImportCompanies]
END
GO


/**************************************************************************8*/
-- CREATE THE ETL IMPORT PROC
/**************************************************************************8*/
CREATE PROCEDURE [dbo].[ImportCompanies] (	@data dbo.Company_Import_tt READONLY)
AS
BEGIN
	SET NOCOUNT ON;

	INSERT INTO [dbo].[Company_Import]
		   ([CompanyId], [StartDate], [EndDate], [LegalName], [DBAName]
		   ,[ChangeDate], [UserId])
	SELECT [CompanyId], [StartDate], [EndDate], [LegalName], [DBAName]
		   ,[ChangeDate], [UserId]
	FROM @data d

END
GO

