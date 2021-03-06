USE [DMIS]
GO
/****** Object:  Table [dbo].[CMIS_OFF_HD]    Script Date: 01/29/2010 14:10:42 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
SET ANSI_PADDING ON
GO
CREATE TABLE [dbo].[CMIS_OFF_HD](
	[ORDER] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[OR_NUM] [nvarchar](10) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[OR_DATE] [datetime] NULL,
	[OR_AMT] [decimal](18, 2) NULL,
	[DISCOUNT] [decimal](18, 2) NULL,
	[TAX] [decimal](18, 2) NULL,
	[CONSUMED] [decimal](18, 2) NULL,
	[BAYADAMT] [decimal](18, 2) NULL,
	[SUKLI] [decimal](18, 2) NULL,
	[CASHAMOUNT] [decimal](18, 2) NULL,
	[CHKAMOUNT] [decimal](18, 2) NULL,
	[CARDAMOUNT] [decimal](18, 2) NULL,
	[CUSCDE] [nvarchar](10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CUSNAME] [nvarchar](250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ADJUSTED] [bit] NULL CONSTRAINT [DF_CMIS_Off_Hd_ADJUSTED]  DEFAULT ((0)),
	[VARIOUS] [bit] NULL CONSTRAINT [DF_CMIS_Off_Hd_VARIOUS]  DEFAULT ((0)),
	[PAIDNA] [bit] NULL CONSTRAINT [DF_CMIS_Off_Hd_PAIDNA]  DEFAULT ((0)),
	[CANCEL] [bit] NULL CONSTRAINT [DF_CMIS_Off_Hd_CANCEL]  DEFAULT ((0)),
	[FORCE] [bit] NULL CONSTRAINT [DF_CMIS_Off_Hd_FORCE]  DEFAULT ((0)),
	[DEPOSIT] [bit] NULL CONSTRAINT [DF_CMIS_Off_Hd_DEPOSIT]  DEFAULT ((0)),
	[BOUNCE] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[TOF] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BANKCODE] [nvarchar](10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BANKBRANCH] [nvarchar](100) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[TSEKE] [nvarchar](30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CARDBNKCDE] [nvarchar](30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CARDNUMBER] [nvarchar](30) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[TSEKLASE] [nvarchar](3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CUTDATE] [datetime] NULL,
	[CHECKDATE] [datetime] NULL,
	[CARDDATE] [char](10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DATECANCEL] [datetime] NULL,
	[TIMECANCEL] [nvarchar](12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[WHOCANCEL] [nvarchar](5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DATECREATE] [datetime] NULL,
	[TIMECREATE] [nvarchar](12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[WHOCREATE] [nvarchar](5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PRINTED] [bit] NULL,
	[VALID] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[STATUS] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NOT NULL,
	[VAT] [int] NULL,
	[REFERENCENO] [nvarchar](8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PAIDBY] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[BANK] [nvarchar](6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DEPOSIT1] [bit] NULL CONSTRAINT [DF_CMIS_Off_Hd_DEPOSIT1]  DEFAULT ((0)),
	[DEPOSIT2] [bit] NULL CONSTRAINT [DF_CMIS_Off_Hd_DEPOSIT2]  DEFAULT ((0)),
	[DEPOSIT3] [bit] NULL CONSTRAINT [DF_CMIS_Off_Hd_DEPOSIT3]  DEFAULT ((0)),
	[TOF1] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[TOF2] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[TOF3] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ID] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]

GO
SET ANSI_PADDING OFF