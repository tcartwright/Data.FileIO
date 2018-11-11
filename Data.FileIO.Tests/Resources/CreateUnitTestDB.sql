
USE [master]
GO--

IF DB_ID('UnitTests') IS NOT NULL BEGIN
	ALTER DATABASE [UnitTests] SET SINGLE_USER WITH ROLLBACK IMMEDIATE; 
	DROP DATABASE [UnitTests]
END
GO--

CREATE DATABASE [UnitTests]
GO--

USE [UnitTests]
GO--


IF OBJECT_ID('[dbo].[Import_Employees]') IS NOT NULL DROP PROCEDURE [dbo].[Import_Employees]
GO--

IF TYPE_ID('[dbo].[EmployeeData]') IS NOT NULL DROP TYPE [dbo].[EmployeeData]
GO--

IF OBJECT_ID('[dbo].[Employees]') IS NOT NULL DROP TABLE [dbo].[Employees]
GO--

SET ANSI_NULLS ON
GO--

SET QUOTED_IDENTIFIER ON
GO--

SET ANSI_PADDING ON
GO--

CREATE TABLE [dbo].[Employees](
	[EmployeeId] [int] IDENTITY(1,1) NOT NULL,
	[SSN] [varchar](50) NOT NULL,
	[Name] [varchar](255) NOT NULL,
	[Errors] varchar(8000) NULL,
 CONSTRAINT [PK_Employees] PRIMARY KEY CLUSTERED 
(
	[EmployeeId] ASC
)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON) ON [PRIMARY]
) ON [PRIMARY]

GO--

SET ANSI_PADDING OFF
GO--

CREATE TYPE [dbo].[EmployeeData] AS TABLE(
	[SSN] varchar(50) NOT NULL,
	[Name] varchar(255) NOT NULL,
	[Errors] varchar(8000) NULL
);
GO--


CREATE PROCEDURE [dbo].[Import_Employees] (@data dbo.EmployeeData READONLY)
AS
BEGIN
	SET NOCOUNT ON;
	INSERT INTO dbo.Employees(SSN, Name, Errors)
	SELECT d.SSN, d.Name, d.Errors 
	FROM @data d
END
GO--