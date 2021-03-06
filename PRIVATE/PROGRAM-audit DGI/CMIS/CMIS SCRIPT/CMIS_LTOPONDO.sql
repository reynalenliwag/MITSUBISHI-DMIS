USE [DMIS]
GO
/****** Object:  Table [dbo].[CMIS_LTOPONDO]    Script Date: 01/29/2010 14:31:37 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CMIS_LTOPONDO](
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
	[LIQUIDATED] [bit] NULL CONSTRAINT [DF_CMIS_LTOPondo_LIQUIDATED]  DEFAULT ((0)),
	[LIQUID] [bit] NULL CONSTRAINT [DF_CMIS_LTOPondo_LIQUID]  DEFAULT ((0)),
	[REPLENISH] [bit] NULL CONSTRAINT [DF_CMIS_LTOPondo_REPLENISH]  DEFAULT ((0)),
	[AR] [bit] NULL CONSTRAINT [DF_CMIS_LTOPONDO_AR]  DEFAULT ((0)),
	[PARTIAL] [bit] NULL CONSTRAINT [DF_CMIS_LTOPondo_PARTIAL]  DEFAULT ((0)),
	[TAG] [bit] NULL CONSTRAINT [DF_CMIS_LTOPondo_TAG]  DEFAULT ((0)),
	[WHOCREATE] [nvarchar](5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ORIGINAL] [decimal](18, 2) NULL,
	[ID] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
