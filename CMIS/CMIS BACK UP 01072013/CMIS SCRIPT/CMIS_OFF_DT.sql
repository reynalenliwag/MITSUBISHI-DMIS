USE [DMIS]
GO
/****** Object:  Table [dbo].[CMIS_OFF_DT]    Script Date: 01/29/2010 14:14:32 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CMIS_OFF_DT](
	[CUTDATE] [datetime] NULL,
	[ORDER] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[OR_NUM] [nvarchar](10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[SUB_OR] [nvarchar](16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[TRANTYPE] [nvarchar](3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[REFERENCE] [nvarchar](250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[INVOICETYPE] [nvarchar](3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[INVOICENO] [nvarchar](11) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CUSCDE] [nvarchar](8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DESCRIPT] [nvarchar](250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BALANCE] [decimal](18, 2) NULL,
	[AMOUNT] [decimal](18, 2) NULL,
	[DOCDTE] [datetime] NULL,
	[ORDATE] [datetime] NULL,
	[PAYMENT] [decimal](18, 2) NULL,
	[DISCOUNT] [decimal](18, 2) NULL,
	[TAX] [decimal](18, 2) NULL,
	[TAX2] [decimal](18, 2) NULL,
	[CONSUMED] [decimal](18, 2) NULL,
	[PRINTED1] [bit] NULL,
	[PRINTED2] [bit] NULL,
	[PRINTED3] [bit] NULL,
	[PRINTED4] [bit] NULL,
	[PAIDNA] [bit] NULL,
	[BOUNCE] [bit] NULL,
	[CANCEL] [bit] NULL,
	[PAIDFOR] [nvarchar](6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BRANCH] [nvarchar](3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[LOCATION] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[OVER] [decimal](18, 2) NULL,
	[PARTYPAY] [bit] NULL,
	[INSURANCE] [nvarchar](3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ORIGINAL_D] [datetime] NULL,
	[VALID] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[STATUS] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[EXPORTED] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CONFIRMED] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[REPORTCODE] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ACCOUNT_CD] [nvarchar](17) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[VAT] [int] NULL,
	[REFERENCENO] [nvarchar](8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ID] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
