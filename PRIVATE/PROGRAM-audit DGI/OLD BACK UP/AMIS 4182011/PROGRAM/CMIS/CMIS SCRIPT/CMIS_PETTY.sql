USE [DMIS]
GO
/****** Object:  Table [dbo].[CMIS_PETTY]    Script Date: 01/29/2010 14:33:32 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CMIS_PETTY](
	[CUTDATE] [datetime] NULL,
	[EMPLOYEE] [nvarchar](20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PETTY_CODE] [nvarchar](3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PETTY_TYPE] [nvarchar](10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ACCOUNT_CD] [nvarchar](17) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PCF_NUMBER] [nvarchar](13) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[PETTY_DATE] [datetime] NULL,
	[PETTY_CASH] [decimal](18, 2) NULL,
	[PARTICULARS] [nvarchar](250) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DATECREATE] [datetime] NULL,
	[TIMECREATE] [nvarchar](12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[LIQUIDTYPE] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[LIQ_AMT] [decimal](18, 2) NULL,
	[LIQ_DATE] [datetime] NULL,
	[LIQUIDATED] [bit] NULL,
	[LIQUID] [bit] NULL,
	[REPLENISH] [bit] NULL CONSTRAINT [DF_CMIS_Petty_REPLENISH]  DEFAULT ((0)),
	[REPL_NO] [nvarchar](6) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[AR] [bit] NULL,
	[PARTIAL] [bit] NULL,
	[TAG] [bit] NULL,
	[WHOCREATE] [nvarchar](5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ORIGINAL] [decimal](18, 2) NULL,
	[ID] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
