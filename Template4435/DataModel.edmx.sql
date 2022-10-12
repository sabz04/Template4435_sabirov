
-- --------------------------------------------------
-- Entity Designer DDL Script for SQL Server 2005, 2008, 2012 and Azure
-- --------------------------------------------------
-- Date Created: 10/12/2022 10:56:35
-- Generated from EDMX file: C:\Users\zulfa\OneDrive\Рабочий стол\лабы исрпо\Template4435_sabirov\Template4435\DataModel.edmx
-- --------------------------------------------------

SET QUOTED_IDENTIFIER OFF;
GO
USE [C:\USERS\ZULFA\ONEDRIVE\ДОКУМЕНТЫ\EXCELDATA.MDF];
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
    [CodeOrder] nvarchar(max)  NOT NULL,
    [CreateDate] nvarchar(max)  NOT NULL,
    [CreateTime] nvarchar(max)  NOT NULL,
    [CodeClient] nvarchar(max)  NOT NULL,
    [Services] nvarchar(max)  NOT NULL,
    [Status] nvarchar(max)  NOT NULL,
    [ClosedDate] nvarchar(max)  NOT NULL,
    [ProkatTime] nvarchar(max)  NOT NULL
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