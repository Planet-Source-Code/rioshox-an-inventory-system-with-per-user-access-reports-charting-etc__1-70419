IF EXISTS (SELECT name FROM master.dbo.sysdatabases WHERE name = N'RFM')
	DROP DATABASE [RFM]
GO

CREATE DATABASE [RFM]  ON (NAME = N'RFM_dat', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL\data\RFM.mdf' , SIZE = 6, FILEGROWTH = 10%) LOG ON (NAME = N'RFM_log', FILENAME = N'C:\Program Files\Microsoft SQL Server\MSSQL\data\RFM.ldf' , SIZE = 2, FILEGROWTH = 10%)
 COLLATE SQL_Latin1_General_CP1_CI_AS
GO

exec sp_dboption N'RFM', N'autoclose', N'true'
GO

exec sp_dboption N'RFM', N'bulkcopy', N'true'
GO

exec sp_dboption N'RFM', N'trunc. log', N'true'
GO

exec sp_dboption N'RFM', N'torn page detection', N'true'
GO

exec sp_dboption N'RFM', N'read only', N'false'
GO

exec sp_dboption N'RFM', N'dbo use', N'false'
GO

exec sp_dboption N'RFM', N'single', N'false'
GO

exec sp_dboption N'RFM', N'autoshrink', N'true'
GO

exec sp_dboption N'RFM', N'ANSI null default', N'false'
GO

exec sp_dboption N'RFM', N'recursive triggers', N'false'
GO

exec sp_dboption N'RFM', N'ANSI nulls', N'false'
GO

exec sp_dboption N'RFM', N'concat null yields null', N'false'
GO

exec sp_dboption N'RFM', N'cursor close on commit', N'false'
GO

exec sp_dboption N'RFM', N'default to local cursor', N'false'
GO

exec sp_dboption N'RFM', N'quoted identifier', N'false'
GO

exec sp_dboption N'RFM', N'ANSI warnings', N'false'
GO

exec sp_dboption N'RFM', N'auto create statistics', N'true'
GO

exec sp_dboption N'RFM', N'auto update statistics', N'true'
GO

if( (@@microsoftversion / power(2, 24) = 8) and (@@microsoftversion & 0xffff >= 724) )
	exec sp_dboption N'RFM', N'db chaining', N'false'
GO

use [RFM]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblReturnItems_tblCustomers]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblReturnItems] DROP CONSTRAINT FK_tblReturnItems_tblCustomers
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblTransactions_tblCustomers]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblTransactions] DROP CONSTRAINT FK_tblTransactions_tblCustomers
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblPosted_tblEmployees]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblPosted] DROP CONSTRAINT FK_tblPosted_tblEmployees
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblSystemUsers_tblEmployees]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblSystemUsers] DROP CONSTRAINT FK_tblSystemUsers_tblEmployees
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblTransactions_tblEmployees]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblTransactions] DROP CONSTRAINT FK_tblTransactions_tblEmployees
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblProducts_tblPackagingTypes]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblProducts] DROP CONSTRAINT FK_tblProducts_tblPackagingTypes
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblProducts_tblProductTypes]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblProducts] DROP CONSTRAINT FK_tblProducts_tblProductTypes
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblBasketDetails_tblProducts]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblBasketDetails] DROP CONSTRAINT FK_tblBasketDetails_tblProducts
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblReturnItemDetails_tblProducts]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblReturnItemDetails] DROP CONSTRAINT FK_tblReturnItemDetails_tblProducts
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblTransactionDetails_tblProducts]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblTransactionDetails] DROP CONSTRAINT FK_tblTransactionDetails_tblProducts
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblReturnItemDetails_tblReturnItems]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblReturnItemDetails] DROP CONSTRAINT FK_tblReturnItemDetails_tblReturnItems
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblPosted_tblTransactions]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblPosted] DROP CONSTRAINT FK_tblPosted_tblTransactions
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[FK_tblTransactionDetails_tblTransactions]') and OBJECTPROPERTY(id, N'IsForeignKey') = 1)
ALTER TABLE [dbo].[tblTransactionDetails] DROP CONSTRAINT FK_tblTransactionDetails_tblTransactions
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[DailySales]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[DailySales]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[GetEmployeeName]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[GetEmployeeName]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[MonthlySales]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[MonthlySales]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[YearlySales]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[YearlySales]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[YearlySalesRanged]') and xtype in (N'FN', N'IF', N'TF'))
drop function [dbo].[YearlySalesRanged]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spChartData]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spChartData]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spChartDataMonthly]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spChartDataMonthly]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spChartDataMonthlyCompare]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spChartDataMonthlyCompare]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spChartDataYear]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spChartDataYear]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spChartDataYearCompare]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spChartDataYearCompare]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spForShipping]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spForShipping]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spPendingSO]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spPendingSO]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spPostedSO]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spPostedSO]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[spTransactionDetails]') and OBJECTPROPERTY(id, N'IsProcedure') = 1)
drop procedure [dbo].[spTransactionDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryBasket]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryBasket]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryDailyProductSales]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryDailyProductSales]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryForShipping]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryForShipping]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryInvoice]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryInvoice]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryMontlyProductSales]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryMontlyProductSales]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryPendingSO]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryPendingSO]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryTransactionDetails]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryTransactionDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qrySearchProduct]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qrySearchProduct]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryTransactions]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryTransactions]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryDailySales]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryDailySales]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryMonthlySales]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryMonthlySales]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qrySearchCustomer]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qrySearchCustomer]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qrySearchEmployee]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qrySearchEmployee]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qrySearchPackaging]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qrySearchPackaging]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qrySearchProductTypes]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qrySearchProductTypes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[qryYearlySales]') and OBJECTPROPERTY(id, N'IsView') = 1)
drop view [dbo].[qryYearlySales]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblAuditTrail]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblAuditTrail]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblBasketDetails]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblBasketDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblCustomers]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblCustomers]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblEmployees]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblEmployees]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblPackagingTypes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblPackagingTypes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblPosted]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblPosted]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblProductTypes]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblProductTypes]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblProducts]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblProducts]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblReturnItemDetails]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblReturnItemDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblReturnItems]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblReturnItems]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblSystemUsers]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblSystemUsers]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblTransactionDetails]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblTransactionDetails]
GO

if exists (select * from dbo.sysobjects where id = object_id(N'[dbo].[tblTransactions]') and OBJECTPROPERTY(id, N'IsUserTable') = 1)
drop table [dbo].[tblTransactions]
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO





CREATE FUNCTION DailySales ()
RETURNS @SalesByDate TABLE
   (
    TransactionID  bigint,
   OrderDate       datetime,
    CustomerID 	   varchar(20)	,    
    Companyname   varchar(50),
    ContactName   varchar(30),
    TransactionTotal money
   )
AS
BEGIN
   INSERT @SalesByDate
        SELECT 
		dbo.tblTransactions.TransactionID, 
		dbo.tblTransactions.ORderDAte, 
		dbo.tblTransactions.CustomerID, 
		dbo.tblCustomers.CompanyName, 
		dbo.tblCustomers.ContactName, 
		SUM((dbo.tblTransactionDetails.UnitPrice * dbo.tblTransactionDetails.Quantity)     * (1 - dbo.tblTransactionDetails.Discount) / 100 * 100) AS TransactionTotal
	FROM 
		dbo.tblTransactions 
	INNER JOIN    
		dbo.tblTransactionDetails 
	ON 
		dbo.tblTransactions.TransactionID = dbo.tblTransactionDetails.TransactionID
	INNER JOIN
		dbo.tblCustomers 
	ON 
		dbo.tblTransactions.CustomerID = dbo.tblCustomers.CustomerID
	GROUP BY 
		dbo.tblTransactions.TransactionID, 
		dbo.tblTransactions.CustomerID, 
		dbo.tblCustomers.CompanyName, 
		dbo.tblCustomers.ContactName,
		dbo.tblTransactions.OrderDate
	ORDER BY
		dbo.tblTransactions.OrderDate 

   RETURN
END








GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE FUNCTION GetEmployeeName (@EmployeeID varchar(150))
RETURNS varchar(150)
AS
BEGIN
	DECLARE @EmployeeName varchar(150)

	SET @EmployeeName =

	(SELECT 		
		a.EmployeeName 
	FROM 
		(
			SELECT 
				EmployeeID, 
			        LastName + ', ' + FirstName + ' ' + LEFT(MiddleName, 1) + '.' AS EmployeeName 
			FROM 
				dbo.tblEmployees
		) a
	WHERE 
		a.EmployeeID = @EmployeeID)
	 
	RETURN @EmployeeName	
END


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO



CREATE  FUNCTION MonthlySales ()
RETURNS @SalesByMonth TABLE
   (
    OrderDate 	   smalldatetime,
    DailyTotal     money,    
    DYear     	   bigint,  	
    DMonth         int
   )
AS
BEGIN
   INSERT @SalesByMonth
        SELECT 
			dbo.tblTransactions.OrderDate, 
			SUM(dbo.tblTransactionDetails.UnitPrice * dbo.tblTransactionDetails.Quantity) AS DailyTotal, 
			{ fn YEAR(dbo.tblTransactions.OrderDate) } AS DYear, 
			{ fn MONTH(dbo.tblTransactions.OrderDate)} AS dmonth
		FROM 
			dbo.tblTransactions 
		INNER JOIN
			dbo.tblTransactionDetails 
		ON 
			dbo.tblTransactions.TransactionID = dbo.tblTransactionDetails.TransactionID
		GROUP BY 
			dbo.tblTransactions.OrderDate, 
			{ fn YEAR(dbo.tblTransactions.OrderDate) }, 
			{ fn MONTH(dbo.tblTransactions.OrderDate) }

		ORDER BY 		
			{ fn MONTH(dbo.tblTransactions.OrderDate)},
			{ fn YEAR(dbo.tblTransactions.OrderDate) } 

   RETURN
END






GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE FUNCTION YearlySales ()
RETURNS @SalesByYear TABLE
   (
    MonthlyTotal   money,    
    SalesYear      bigint,  	
    SalesMonth     int
   )
AS
BEGIN
   INSERT @SalesByYear
        SELECT 			
			SUM(dbo.tblTransactionDetails.UnitPrice * dbo.tblTransactionDetails.Quantity) AS MonthlyTotal, 
			{ fn YEAR(dbo.tblTransactions.OrderDate) } AS SalesYear, 
			{ fn MONTH(dbo.tblTransactions.OrderDate)} AS SalesMonth
		FROM 
			dbo.tblTransactions 
		INNER JOIN
			dbo.tblTransactionDetails 
		ON 
			dbo.tblTransactions.TransactionID = dbo.tblTransactionDetails.TransactionID
		GROUP BY
			{ fn YEAR(dbo.tblTransactions.OrderDate) }, 
			{ fn MONTH(dbo.tblTransactions.OrderDate) }		
		ORDER BY 		
			{ fn YEAR(dbo.tblTransactions.OrderDate) }, 
			{ fn MONTH(dbo.tblTransactions.OrderDate)}

   RETURN
END





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS OFF 
GO


CREATE FUNCTION YearlySalesRanged (@StartYear bigint,@EndYEar bigint)
RETURNS @SalesByYear TABLE
   (
    MonthlyTotal   money,    
    SalesYear      bigint,  	
    SalesMonth     int
   )
AS
BEGIN
   INSERT @SalesByYear
        SELECT 			
			SUM(dbo.tblTransactionDetails.UnitPrice * dbo.tblTransactionDetails.Quantity) AS MonthlyTotal, 
			{ fn YEAR(dbo.tblTransactions.OrderDate) } AS SalesYear, 
			{ fn MONTH(dbo.tblTransactions.OrderDate)} AS SalesMonth
		FROM 
			dbo.tblTransactions 
		INNER JOIN
			dbo.tblTransactionDetails 
		ON 
			dbo.tblTransactions.TransactionID = dbo.tblTransactionDetails.TransactionID
		GROUP BY
			{ fn YEAR(dbo.tblTransactions.OrderDate) }, 
			{ fn MONTH(dbo.tblTransactions.OrderDate) }
		HAVING 
			({ fn YEAR(dbo.tblTransactions.OrderDate) } >= @StartYear 
		AND 
			{ fn YEAR(dbo.tblTransactions.OrderDate) } <=@EndYEar)
		ORDER BY 		
			{ fn YEAR(dbo.tblTransactions.OrderDate) }, 
			{ fn MONTH(dbo.tblTransactions.OrderDate)}

   RETURN
END





GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

CREATE TABLE [dbo].[tblAuditTrail] (
	[AuditID] [int] IDENTITY (1, 1) NOT NULL ,
	[ActionText] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EmployeeID] [int] NULL ,
	[ActionDate] [smalldatetime] NULL ,
	[ActionTime] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblBasketDetails] (
	[SessionID] [int] NOT NULL ,
	[ProductID] [int] NOT NULL ,
	[Quantity] [float] NOT NULL ,
	[UnitPrice] [money] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblCustomers] (
	[CustomerID] [nvarchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[CompanyName] [nvarchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ContactName] [nvarchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Address] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[City] [nvarchar] (40) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Region] [nvarchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PostalCode] [nvarchar] (5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Phone] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SMSNumber] [nvarchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[SMSIPin] [nvarchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[CreditLimit] [money] NULL ,
	[MinimumPurchase] [money] NULL ,
	[EmailAddress] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblEmployees] (
	[EmployeeID] [int] IDENTITY (1, 1) NOT NULL ,
	[LastName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[FirstName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[MiddleName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Gender] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Birthdate] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Department] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[Position] [nvarchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[HireDate] [smalldatetime] NULL ,
	[Address] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[City] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ZIP] [nvarchar] (4) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ContactNo] [nvarchar] (30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EmailAddress] [nvarchar] (100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[PicturePath] [nvarchar] (25) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblPackagingTypes] (
	[PackagingTypeID] [int] IDENTITY (1, 1) NOT NULL ,
	[PackagingType] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblPosted] (
	[TransactionID] [bigint] NOT NULL ,
	[EmployeeID] [int] NOT NULL ,
	[PostDate] [smalldatetime] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblProductTypes] (
	[ProductTypeID] [int] IDENTITY (1, 1) NOT NULL ,
	[ProductType] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblProducts] (
	[ProductID] [int] IDENTITY (1, 1) NOT NULL ,
	[ProductName] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[ProductTypeID] [int] NULL ,
	[PackagingTypeID] [int] NULL ,
	[QuantityPerUnit] [nvarchar] (255) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[UnitPrice] [money] NULL ,
	[UnitsInStock] [float] NULL ,
	[UnitsOnOrder] [float] NULL ,
	[ReorderLevel] [smallint] NULL ,
	[Discontinued] [bit] NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblReturnItemDetails] (
	[ReturnID] [bigint] NOT NULL ,
	[ProductID] [int] NOT NULL ,
	[Quantity] [decimal](18, 0) NOT NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblReturnItems] (
	[ReturnID] [bigint] IDENTITY (1, 1) NOT NULL ,
	[CustomerID] [nvarchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[ReturnDate] [smalldatetime] NULL ,
	[ReferenceID] [int] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblSystemUsers] (
	[UserName] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[Password] [nvarchar] (50) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[EmployeeID] [int] NOT NULL ,
	[IsAdmin] [bit] NOT NULL ,
	[Products] [bit] NULL ,
	[Customer] [bit] NULL ,
	[UserAccess] [bit] NULL ,
	[ApproveSO] [bit] NULL ,
	[ApproveShip] [bit] NULL ,
	[BackupData] [bit] NULL ,
	[ReturnItem] [bit] NULL ,
	[ManualSO] [bit] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblTransactionDetails] (
	[TransactionID] [bigint] NOT NULL ,
	[ProductID] [int] NOT NULL ,
	[UnitPrice] [money] NOT NULL ,
	[Quantity] [int] NOT NULL ,
	[Discount] [real] NULL 
) ON [PRIMARY]
GO

CREATE TABLE [dbo].[tblTransactions] (
	[TransactionID] [bigint] IDENTITY (1, 1) NOT NULL ,
	[CustomerID] [nvarchar] (20) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL ,
	[OrderDate] [smalldatetime] NOT NULL ,
	[RequiredDate] [smalldatetime] NULL ,
	[ShippedDate] [smalldatetime] NULL ,
	[OrderSource] [nvarchar] (10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL ,
	[EmployeeID] [int] NULL ,
	[isPickup] [bit] NULL 
) ON [PRIMARY]
GO

ALTER TABLE [dbo].[tblReturnItemDetails] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblReturnItemDetails] PRIMARY KEY  CLUSTERED 
	(
		[ReturnID],
		[ProductID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblReturnItems] WITH NOCHECK ADD 
	CONSTRAINT [PK_tblReturnItems] PRIMARY KEY  CLUSTERED 
	(
		[ReturnID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblAuditTrail] ADD 
	CONSTRAINT [PK_tblAuditTrail] PRIMARY KEY  NONCLUSTERED 
	(
		[AuditID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblCustomers] ADD 
	CONSTRAINT [PK_tblCustomers] PRIMARY KEY  NONCLUSTERED 
	(
		[CustomerID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblEmployees] ADD 
	CONSTRAINT [PK_tblEmployees] PRIMARY KEY  NONCLUSTERED 
	(
		[EmployeeID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblPackagingTypes] ADD 
	CONSTRAINT [PK_tblPackagingTypes] PRIMARY KEY  NONCLUSTERED 
	(
		[PackagingTypeID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblPosted] ADD 
	CONSTRAINT [PK_tblPosted] PRIMARY KEY  NONCLUSTERED 
	(
		[TransactionID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblProductTypes] ADD 
	CONSTRAINT [PK_tblProductTypes] PRIMARY KEY  NONCLUSTERED 
	(
		[ProductTypeID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblProducts] ADD 
	CONSTRAINT [PK_tblProducts] PRIMARY KEY  NONCLUSTERED 
	(
		[ProductID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblSystemUsers] ADD 
	CONSTRAINT [DF_tblSystemUsers_Products] DEFAULT (0) FOR [Products],
	CONSTRAINT [DF_tblSystemUsers_Customer] DEFAULT (0) FOR [Customer],
	CONSTRAINT [DF_tblSystemUsers_UserAccess] DEFAULT (0) FOR [UserAccess],
	CONSTRAINT [DF_tblSystemUsers_ApproveSO] DEFAULT (0) FOR [ApproveSO],
	CONSTRAINT [DF_tblSystemUsers_ApproveShip] DEFAULT (0) FOR [ApproveShip],
	CONSTRAINT [DF_tblSystemUsers_BackupData] DEFAULT (0) FOR [BackupData],
	CONSTRAINT [PK_tblSystemUsers] PRIMARY KEY  NONCLUSTERED 
	(
		[EmployeeID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblTransactionDetails] ADD 
	CONSTRAINT [DF_tblTransactionDetails_Discount] DEFAULT (0) FOR [Discount],
	CONSTRAINT [PK_tblTransactionDetails] PRIMARY KEY  NONCLUSTERED 
	(
		[TransactionID],
		[ProductID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblTransactions] ADD 
	CONSTRAINT [PK_tblTransactions] PRIMARY KEY  NONCLUSTERED 
	(
		[TransactionID]
	)  ON [PRIMARY] 
GO

ALTER TABLE [dbo].[tblBasketDetails] ADD 
	CONSTRAINT [FK_tblBasketDetails_tblProducts] FOREIGN KEY 
	(
		[ProductID]
	) REFERENCES [dbo].[tblProducts] (
		[ProductID]
	)
GO

ALTER TABLE [dbo].[tblPosted] ADD 
	CONSTRAINT [FK_tblPosted_tblEmployees] FOREIGN KEY 
	(
		[EmployeeID]
	) REFERENCES [dbo].[tblEmployees] (
		[EmployeeID]
	),
	CONSTRAINT [FK_tblPosted_tblTransactions] FOREIGN KEY 
	(
		[TransactionID]
	) REFERENCES [dbo].[tblTransactions] (
		[TransactionID]
	)
GO

ALTER TABLE [dbo].[tblProducts] ADD 
	CONSTRAINT [FK_tblProducts_tblPackagingTypes] FOREIGN KEY 
	(
		[PackagingTypeID]
	) REFERENCES [dbo].[tblPackagingTypes] (
		[PackagingTypeID]
	),
	CONSTRAINT [FK_tblProducts_tblProductTypes] FOREIGN KEY 
	(
		[ProductTypeID]
	) REFERENCES [dbo].[tblProductTypes] (
		[ProductTypeID]
	)
GO

ALTER TABLE [dbo].[tblReturnItemDetails] ADD 
	CONSTRAINT [FK_tblReturnItemDetails_tblProducts] FOREIGN KEY 
	(
		[ProductID]
	) REFERENCES [dbo].[tblProducts] (
		[ProductID]
	) ON UPDATE CASCADE ,
	CONSTRAINT [FK_tblReturnItemDetails_tblReturnItems] FOREIGN KEY 
	(
		[ReturnID]
	) REFERENCES [dbo].[tblReturnItems] (
		[ReturnID]
	) ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[tblReturnItems] ADD 
	CONSTRAINT [FK_tblReturnItems_tblCustomers] FOREIGN KEY 
	(
		[CustomerID]
	) REFERENCES [dbo].[tblCustomers] (
		[CustomerID]
	) ON UPDATE CASCADE 
GO

ALTER TABLE [dbo].[tblSystemUsers] ADD 
	CONSTRAINT [FK_tblSystemUsers_tblEmployees] FOREIGN KEY 
	(
		[EmployeeID]
	) REFERENCES [dbo].[tblEmployees] (
		[EmployeeID]
	)
GO

ALTER TABLE [dbo].[tblTransactionDetails] ADD 
	CONSTRAINT [FK_tblTransactionDetails_tblProducts] FOREIGN KEY 
	(
		[ProductID]
	) REFERENCES [dbo].[tblProducts] (
		[ProductID]
	),
	CONSTRAINT [FK_tblTransactionDetails_tblTransactions] FOREIGN KEY 
	(
		[TransactionID]
	) REFERENCES [dbo].[tblTransactions] (
		[TransactionID]
	)
GO

ALTER TABLE [dbo].[tblTransactions] ADD 
	CONSTRAINT [FK_tblTransactions_tblCustomers] FOREIGN KEY 
	(
		[CustomerID]
	) REFERENCES [dbo].[tblCustomers] (
		[CustomerID]
	),
	CONSTRAINT [FK_tblTransactions_tblEmployees] FOREIGN KEY 
	(
		[EmployeeID]
	) REFERENCES [dbo].[tblEmployees] (
		[EmployeeID]
	)
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO




CREATE VIEW qryDailySales as SELECT * FROM dbo.DailySales() Go


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO



CREaTE VIEW qryMonthlySales 
AS
SELECT * FROM MonthlySales()


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW dbo.qrySearchCustomer
AS
SELECT tblCustomers.CustomerID, tblCustomers.CompanyName, 
    tblCustomers.ContactName, tblCustomers.Phone
FROM tblCustomers


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW dbo.qrySearchEmployee
AS
SELECT [tblEmployees].[EmployeeID], 
    [LastName] + ', ' + [FirstName] + ' ' + LEFT([MiddleName], 1) 
    + '.' AS EmployeeName, [tblEmployees].[Position]
FROM tblEmployees


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW dbo.qrySearchPackaging
AS
SELECT tblPackagingTypes.PackagingTypeID, 
    tblPackagingTypes.PackagingType
FROM tblPackagingTypes


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW dbo.qrySearchProductTypes
AS
SELECT tblProductTypes.ProductTypeID, 
    tblProductTypes.ProductType
FROM tblProductTypes


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW qryYearlySales as 
	SELECT * FROM dbo.YearlySales()


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE  VIEW dbo.qrySearchProduct
AS
SELECT dbo.tblProducts.ProductID, dbo.tblProducts.ProductName, 
    dbo.tblPackagingTypes.PackagingType, 
    dbo.tblProducts.UnitsInStock, dbo.tblProducts.UnitsOnOrder, 
    dbo.tblProducts.QuantityPerUnit, dbo.tblProducts.UnitPrice, 
    dbo.tblProducts.ReorderLevel, 
    dbo.tblProductTypes.ProductType
FROM dbo.tblPackagingTypes INNER JOIN
    dbo.tblProductTypes INNER JOIN
    dbo.tblProducts ON 
    dbo.tblProductTypes.ProductTypeID = dbo.tblProducts.ProductTypeID
     ON 
    dbo.tblPackagingTypes.PackagingTypeID = dbo.tblProducts.PackagingTypeID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE  VIEW dbo.qryTransactions
AS
SELECT tblTransactions.CustomerID, ISNULL([CompanyName], 
    [ContactName]) AS Customer, tblCustomers.Phone, 
    tblTransactions.OrderDate, tblTransactions.RequiredDate, 
    tblTransactions.ShippedDate, tblTransactions.OrderSource, 
    tblTransactions.TransactionID, 
    tblTransactions.EmployeeID
FROM tblCustomers INNER JOIN
    tblTransactions ON 
    tblCustomers.CustomerID = tblTransactions.CustomerID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW qryBasket
AS 
	SELECT tblBasketDetails.ProductID, tblProducts.ProductName, tblProductTypes.ProductType, tblPackagingTypes.PackagingType, tblBasketDetails.Quantity, tblBasketDetails.UnitPrice, tblBasketDetails.SessionID
	FROM (tblPackagingTypes INNER JOIN (tblProductTypes INNER JOIN tblProducts ON tblProductTypes.ProductTypeID = tblProducts.ProductTypeID) ON tblPackagingTypes.PackagingTypeID = tblProducts.PackagingTypeID) INNER JOIN tblBasketDetails ON tblProducts.ProductID = tblBasketDetails.ProductID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW qryDailyProductSales 
AS

SELECT 
	dbo.tblTransactions.OrderDate, 
	dbo.tblTransactionDetails.ProductID, 
	dbo.qrySearchProduct.ProductName, 
	dbo.qrySearchProduct.PackagingType, 
	dbo.qrySearchProduct.ProductType, 
	SUM((dbo.tblTransactionDetails.UnitPrice * dbo.tblTransactionDetails.Quantity) * (1 - dbo.tblTransactionDetails.Discount) / 100 * 100) AS ProductSales
FROM 
	dbo.tblTransactions 
INNER JOIN
	dbo.tblTransactionDetails 
ON 
	dbo.tblTransactions.TransactionID = dbo.tblTransactionDetails.TransactionID
RIGHT OUTER JOIN
	dbo.qrySearchProduct 
ON 
	dbo.tblTransactionDetails.ProductID = dbo.qrySearchProduct.ProductID
GROUP BY 
	dbo.tblTransactions.OrderDate, 
	dbo.tblTransactionDetails.ProductID, 
	dbo.qrySearchProduct.ProductName, 
	dbo.qrySearchProduct.PackagingType, 
	dbo.qrySearchProduct.ProductType


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE view qryForShipping  as
SELECT TOP 100 PERCENT  * FROM qryTransactions WHERE ShippedDate Is Null AND EmployeeID <> 0 Order BY OrderDate


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE  VIEW qryInvoice 
AS
SELECT 
	dbo.tblTransactions.TransactionID, 	
	dbo.tblTransactions.OrderDate, 
	dbo.tblTransactions.RequiredDate, 
	dbo.tblTransactions.ShippedDate, 	
	dbo.tblTransactions.CustomerID, 	
	CASE Customers.CompanyName
		WHEN Null THEN Customers.ContactName
		ELSE  Customers.CompanyName
	END CustomerName,	
	Customers.Address + ',' + Customers.City + ' ' + Customers.PostalCode CustAddress,
	Customers.Phone,
	dbo.tblTransactions.OrderSource, 


	dbo.tblTransactionDetails.ProductID, 
	Products.ProductName, 
	Products.PackagingType, 
	Products.ProductType, 
	dbo.tblTransactionDetails.UnitPrice, 
	dbo.tblTransactionDetails.Quantity, 
	dbo.tblTransactionDetails.Discount, 
    	(tblTransactionDetails.UnitPrice*Quantity*(1-Discount)/100) * 100 ExtendedPrice,
    
    
	dbo.GEtEmployeeName(dbo.tblTransactions.EmployeeID) ApprovedBy, 
	dbo.GetEmployeeName(dbo.tblPosted.EmployeeID) AS PostedBy     
    
FROM 
	dbo.tblTransactions 
INNER JOIN
	dbo.tblTransactionDetails 
ON 
    	dbo.tblTransactions.TransactionID = dbo.tblTransactionDetails.TransactionID
INNER JOIN
    	(
		SELECT 
			tblProducts.ProductID, 
			tblProducts.ProductName, 
			tblPackagingTypes.PackagingType, 
			tblProducts.UnitsInStock, 
			tblProducts.UnitsOnOrder, 
			tblProducts.QuantityPerUnit, 
			tblProducts.UnitPrice, 
			tblProducts.ReorderLevel, 
			tblProductTypes.ProductType
		FROM 
			dbo.tblPackagingTypes 
		INNER JOIN
			dbo.tblProductTypes 
		INNER JOIN
			dbo.tblProducts 
		ON 
			dbo.tblProductTypes.ProductTypeID = dbo.tblProducts.ProductTypeID
		ON 
			dbo.tblPackagingTypes.PackagingTypeID = dbo.tblProducts.PackagingTypeID

		) Products 
ON 
    	dbo.tblTransactionDetails.ProductID = Products.ProductID
INNER JOIN
	(
		SELECT 
			CustomerID, 
			CompanyName, 
			ContactName, 
			Address, City, PostalCode,
			Phone
		FROM 	
			dbo.tblCustomers
	)Customers 
ON 
	dbo.tblTransactions.CustomerID = Customers.CustomerID
LEFT OUTER JOIN
    	dbo.tblPosted 
ON 
	dbo.tblTransactions.TransactionID = dbo.tblPosted.TransactionID



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW qryMontlyProductSales 
AS

	SELECT 
		MONTH(tblTransactions.OrderDate) SalesMonth, 
		YEAR(tblTransactions.OrderDate) SalesYear, 
		tblTransactionDetails.ProductID, 
		qrySearchProduct.ProductName, 
		qrySearchProduct.PackagingType, 
		qrySearchProduct.ProductType, 
		SUM((tblTransactionDetails.UnitPrice * tblTransactionDetails.Quantity)* (1 - tblTransactionDetails.Discount) / 100 * 100) AS ProductSales
	FROM 
		dbo.tblTransactions 
	INNER JOIN
		dbo.tblTransactionDetails 
	ON 
		dbo.tblTransactions.TransactionID = dbo.tblTransactionDetails.TransactionID
	RIGHT OUTER JOIN
		dbo.qrySearchProduct 
	ON 
		dbo.tblTransactionDetails.ProductID = dbo.qrySearchProduct.ProductID
	GROUP BY 
		MONTH(dbo.tblTransactions.OrderDate), 
		YEAR(dbo.tblTransactions.OrderDate), 
		dbo.tblTransactionDetails.ProductID, 
		dbo.qrySearchProduct.ProductName, 
		dbo.qrySearchProduct.PackagingType, 
		dbo.qrySearchProduct.ProductType



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE view qryPendingSO as
SELECT TOP 100 PERCENT  * FROM qryTransactions WHERE ShippedDate Is Null AND EmployeeID = 0 Order BY OrderDate


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE VIEW dbo.qryTransactionDetails
AS
SELECT tblTransactionDetails.ProductID, 
    tblTransactionDetails.TransactionID, 
    tblTransactionDetails.UnitPrice, tblTransactionDetails.Quantity, 
    tblTransactionDetails.Discount, 
    qrySearchProduct.ProductName, 
    qrySearchProduct.PackagingType, 
    (dbo.tblTransactionDetails.UnitPrice * dbo.tblTransactionDetails.Quantity
     * (1 - Discount) / 100) * 100 ExtendedPrice, 
    qrySearchProduct.ProductType
FROM dbo.qrySearchProduct INNER JOIN
    dbo.tblTransactionDetails ON 
    dbo.qrySearchProduct.ProductID = dbo.tblTransactionDetails.ProductID


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE spChartData (@StartYear bigint,@EndYEar bigint)
	AS
		SELECT 
			dbo.tblTransactions.OrderDate, 
			SUM(dbo.tblTransactionDetails.UnitPrice * dbo.tblTransactionDetails.Quantity) AS DailyTotal, 
			{ fn YEAR(dbo.tblTransactions.OrderDate) } AS DYear, 
			{ fn MONTH(dbo.tblTransactions.OrderDate)} AS dmonth
		FROM 
			dbo.tblTransactions 
		INNER JOIN
			dbo.tblTransactionDetails 
		ON 
			dbo.tblTransactions.TransactionID = dbo.tblTransactionDetails.TransactionID
		GROUP BY 
			dbo.tblTransactions.OrderDate, 
			{ fn YEAR(dbo.tblTransactions.OrderDate) }, 
			{ fn MONTH(dbo.tblTransactions.OrderDate) }
		HAVING 
			({ fn YEAR(dbo.tblTransactions.OrderDate) } >= @StartYear 
		AND 
			{ fn YEAR(dbo.tblTransactions.OrderDate) } <=@EndYEar)
		ORDER BY 		
			{ fn MONTH(dbo.tblTransactions.OrderDate)},
			{ fn YEAR(dbo.tblTransactions.OrderDate) } 


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE PROCEDURE spChartDataMonthly (@StartYear bigint,@EndYEar bigint)
	AS
		SELECT 			
			SUM(dbo.tblTransactionDetails.UnitPrice * dbo.tblTransactionDetails.Quantity) AS MonthlyTotal, 
			{ fn YEAR(dbo.tblTransactions.OrderDate) } AS DYear, 
			{ fn MONTH(dbo.tblTransactions.OrderDate)} AS dmonth
		FROM 
			dbo.tblTransactions 
		INNER JOIN
			dbo.tblTransactionDetails 
		ON 
			dbo.tblTransactions.TransactionID = dbo.tblTransactionDetails.TransactionID
		GROUP BY
			{ fn YEAR(dbo.tblTransactions.OrderDate) }, 
			{ fn MONTH(dbo.tblTransactions.OrderDate) }
		HAVING 
			({ fn YEAR(dbo.tblTransactions.OrderDate) } >= @StartYear 
		AND 
			{ fn YEAR(dbo.tblTransactions.OrderDate) } <=@EndYEar)
		ORDER BY 		
			{ fn YEAR(dbo.tblTransactions.OrderDate) }, 
			{ fn MONTH(dbo.tblTransactions.OrderDate)}




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE PROCEDURE spChartDataMonthlyCompare (@Year1 bigint,@YEar2 bigint)
	AS
		SELECT 			
			SUM(dbo.tblTransactionDetails.UnitPrice * dbo.tblTransactionDetails.Quantity) AS MonthlyTotal, 
			{ fn YEAR(dbo.tblTransactions.OrderDate) } AS DYear, 
			{ fn MONTH(dbo.tblTransactions.OrderDate)} AS dmonth
		FROM 
			dbo.tblTransactions 
		INNER JOIN
			dbo.tblTransactionDetails 
		ON 
			dbo.tblTransactions.TransactionID = dbo.tblTransactionDetails.TransactionID
		GROUP BY
			{ fn YEAR(dbo.tblTransactions.OrderDate) }, 
			{ fn MONTH(dbo.tblTransactions.OrderDate) }
		HAVING 
			({ fn YEAR(dbo.tblTransactions.OrderDate) } = @Year1 
		OR 
			{ fn YEAR(dbo.tblTransactions.OrderDate) } <=@Year2)
		ORDER BY 		
			{ fn YEAR(dbo.tblTransactions.OrderDate) }, 
			{ fn MONTH(dbo.tblTransactions.OrderDate)}




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE spChartDataYear (@StartYear bigint,@EndYEar bigint)
	AS
		SELECT DISTINCT c.dYear FROM
		(SELECT 
			dbo.tblTransactions.OrderDate, 
			SUM(dbo.tblTransactionDetails.UnitPrice * dbo.tblTransactionDetails.Quantity) AS DailyTotal, 
			{ fn YEAR(dbo.tblTransactions.OrderDate) } AS DYear, 
			{ fn MONTH(dbo.tblTransactions.OrderDate)} AS dmonth
		FROM 
			dbo.tblTransactions 
		INNER JOIN
			dbo.tblTransactionDetails 
		ON 
			dbo.tblTransactions.TransactionID = dbo.tblTransactionDetails.TransactionID
		GROUP BY 
			dbo.tblTransactions.OrderDate, 
			{ fn YEAR(dbo.tblTransactions.OrderDate) }, 
			{ fn MONTH(dbo.tblTransactions.OrderDate) }
		HAVING 
			({ fn YEAR(dbo.tblTransactions.OrderDate) } >= @StartYear 
		AND 
			{ fn YEAR(dbo.tblTransactions.OrderDate) } <=@EndYEar)
		) c
		ORDER BY c.DYear


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE spChartDataYearCompare (@Year1 bigint,@Year2 bigint)
	AS
		SELECT DISTINCT c.dYear FROM
		(SELECT 
			dbo.tblTransactions.OrderDate, 
			SUM(dbo.tblTransactionDetails.UnitPrice * dbo.tblTransactionDetails.Quantity) AS DailyTotal, 
			{ fn YEAR(dbo.tblTransactions.OrderDate) } AS DYear, 
			{ fn MONTH(dbo.tblTransactions.OrderDate)} AS dmonth
		FROM 
			dbo.tblTransactions 
		INNER JOIN
			dbo.tblTransactionDetails 
		ON 
			dbo.tblTransactions.TransactionID = dbo.tblTransactionDetails.TransactionID
		GROUP BY 
			dbo.tblTransactions.OrderDate, 
			{ fn YEAR(dbo.tblTransactions.OrderDate) }, 
			{ fn MONTH(dbo.tblTransactions.OrderDate) }
		HAVING 
			({ fn YEAR(dbo.tblTransactions.OrderDate) } = @Year1 
		OR
			{ fn YEAR(dbo.tblTransactions.OrderDate) } =@Year2)
		) c
		ORDER BY c.DYear


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO


CREATE PROCEDURE spForShipping

AS
	SELECT TransactionDetails.*
 FROM (
	SELECT 
		tblTransactions.CustomerID, 
		ISNULL(tblCustomers.CompanyName, tblCustomers.ContactName) AS Customer, tblCustomers.Phone, 
		tblTransactions.OrderDate, 
		tblTransactions.RequiredDate, 
		tblTransactions.ShippedDate, 
		tblTransactions.OrderSource, 
	 	tblTransactions.TransactionID, 
	        tblTransactions.EmployeeID
	FROM 
		dbo.tblCustomers 
	INNER JOIN
		dbo.tblTransactions 
	ON 
		dbo.tblCustomers.CustomerID = dbo.tblTransactions.CustomerID
	) TransactionDetails
WHERE 
	ShippedDate Is Null 
AND 
	EmployeeID <> 0 
ORDER BY 
	OrderDate


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE  PROCEDURE spPendingSO

AS
	SELECT TransactionDetails.*
 FROM (
	SELECT 
		tblTransactions.CustomerID, 
		ISNULL(tblCustomers.CompanyName, tblCustomers.ContactName) AS Customer, tblCustomers.Phone, 
		tblTransactions.OrderDate, 
		tblTransactions.RequiredDate, 
		tblTransactions.ShippedDate, 
		tblTransactions.OrderSource, 
	 	tblTransactions.TransactionID, 
	        tblTransactions.EmployeeID
	FROM 
		dbo.tblCustomers 
	INNER JOIN
		dbo.tblTransactions 
	ON 
		dbo.tblCustomers.CustomerID = dbo.tblTransactions.CustomerID
	) TransactionDetails
WHERE 
	ShippedDate Is Null 
AND 
	EmployeeID = 0 
ORDER BY 
	OrderDate



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO




CREATE   PROCEDURE spPostedSO @StartDate smalldatetime, @EndDate smalldatetime

AS
	SELECT TransactionDetails.*
 FROM (
	SELECT 
		tblTransactions.CustomerID, 
		ISNULL(tblCustomers.CompanyName, tblCustomers.ContactName) AS Customer, tblCustomers.Phone, 
		tblTransactions.OrderDate, 
		tblTransactions.RequiredDate, 
		tblTransactions.ShippedDate, 
		tblTransactions.OrderSource, 
	 	tblTransactions.TransactionID, 
	        tblTransactions.EmployeeID
	FROM 
		dbo.tblCustomers 
	INNER JOIN
		dbo.tblTransactions 
	ON 
		dbo.tblCustomers.CustomerID = dbo.tblTransactions.CustomerID
	) TransactionDetails
WHERE 
	ShippedDate Is Not Null 
AND 
	EmployeeID <> 0 
AND 
	OrderDate BETWEEN @StartDate AND @EndDate
ORDER BY 
	OrderDate




GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

SET QUOTED_IDENTIFIER ON 
GO
SET ANSI_NULLS ON 
GO



CREATE  PROCEDURE spTransactionDetails @TransactionID bigint
AS
	SELECT 
	    qrySearchProduct.ProductName, 
	    qrySearchProduct.PackagingType, 
	    qrySearchProduct.ProductType, 
	    dbo.tblTransactionDetails.ProductID, 
	    dbo.tblTransactionDetails.UnitPrice, 
	    dbo.tblTransactionDetails.Quantity, 
	    dbo.tblTransactionDetails.Discount,
	   (dbo.tblTransactionDetails.UnitPrice*dbo.tblTransactionDetails.Quantity*(1-Discount)/100)*100 ExtendedPrice	 
	FROM 
	    (
		SELECT 
		    tblProducts.ProductID, 
		    tblProducts.ProductName, 
		    tblPackagingTypes.PackagingType, 
	 	    tblProducts.UnitsInStock, 
		    tblProducts.UnitsOnOrder, 
		    tblProducts.QuantityPerUnit, 
		    tblProducts.UnitPrice, 
		    tblProducts.ReorderLevel, 
		    tblProductTypes.ProductType		    
		FROM 
		    dbo.tblPackagingTypes 
		INNER JOIN
		    dbo.tblProductTypes 
		INNER JOIN
		    dbo.tblProducts ON 
		    dbo.tblProductTypes.ProductTypeID = dbo.tblProducts.ProductTypeID
		ON 
		    dbo.tblPackagingTypes.PackagingTypeID = dbo.tblProducts.PackagingTypeID
	   )qrySearchProduct 
	INNER JOIN
	    dbo.tblTransactionDetails 
	ON 
	    qrySearchProduct.ProductID = dbo.tblTransactionDetails.ProductID
	WHERE
		tblTransactionDetails.TransactionID = @TransactionID 



GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

