
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 10/10/2022 22:54:14
-- Generated from EDMX file: E:\Sabirov2ISRPO\Template4435_sabirov\Template4435\DataModel.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [C:\USERS\GAMER1070\DOCUMENTS\EXCELDB.MDF];
GO
IF SCHEMA_ID(N'dbo') IS NULL EXECUTE(N'CREATE SCHEMA [dbo]');
GO

-- --------------------------------------------------
-- Dropping existing FOREIGN KEY constraints
-- --------------------------------------------------


-- --------------------------------------------------
-- Dropping existing tables
-- --------------------------------------------------

IF OBJECT_ID(N'[dbo].[ExcelDataSet]', 'U') IS NOT NULL
    DROP TABLE [dbo].[ExcelDataSet];
GO

-- --------------------------------------------------
-- Creating all tables
-- --------------------------------------------------

-- Creating table 'ExcelDataSet'
CREATE TABLE [dbo].[ExcelDataSet] (
    [Id] int  NOT NULL,
    [OrderCode] nvarchar(max)  NOT NULL,
    [Date] nvarchar(max)  NOT NULL,
    [Time] nvarchar(max)  NOT NULL,
    [UserCode] nvarchar(max)  NOT NULL,
    [Services] nvarchar(max)  NOT NULL,
    [Status] nvarchar(max)  NOT NULL,
    [DateofClose] nvarchar(max)  NOT NULL,
    [RentalTime] nvarchar(max)  NOT NULL
);
GO

-- --------------------------------------------------
-- Creating all PRIMARY KEY constraints
-- --------------------------------------------------

-- Creating primary key on [Id] in table 'ExcelDataSet'
ALTER TABLE [dbo].[ExcelDataSet]
ADD CONSTRAINT [PK_ExcelDataSet]
    PRIMARY KEY CLUSTERED ([Id] ASC);
GO

-- --------------------------------------------------
-- Creating all FOREIGN KEY constraints
-- --------------------------------------------------

-- --------------------------------------------------
-- Script has ended
-- --------------------------------------------------