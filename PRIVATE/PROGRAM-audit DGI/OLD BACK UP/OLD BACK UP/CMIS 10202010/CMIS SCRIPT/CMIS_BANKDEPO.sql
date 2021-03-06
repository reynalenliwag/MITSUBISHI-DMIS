USE [DMIS]
GO
/****** Object:  Table [dbo].[CMIS_BANKDEPO]    Script Date: 01/29/2010 14:17:53 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CMIS_BANKDEPO](
	[BANKCODE] [nvarchar](8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[TSEKLASE] [nvarchar](3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DEPOSIT] [decimal](18, 2) NULL,
	[DATDEPOSIT] [datetime] NULL,
	[TIMDEPOSIT] [nvarchar](20) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[WHODEPOSIT] [nvarchar](5) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[TYPE] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CUTDATE] [datetime] NULL,
	[INCASHCHK] [bit] NULL,
	[COLLECTCHK] [bit] NULL,
	[P_PAY_CHK] [bit] NULL,
	[L_PAY_CHK] [bit] NULL,
	[U_PAY_CHK] [bit] NULL,
	[A_PAY_CHK] [bit] NULL,
	[PAYMENTAMT] [decimal](18, 2) NULL,
	[INCASHAMT] [decimal](18, 2) NULL,
	[LTOPAYMENT] [decimal](18, 2) NULL,
	[LTOINCASH] [decimal](18, 2) NULL,
	[OR_NUM] [nvarchar](10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DATECREATE] [datetime] NULL,
	[TIMECREATE] [nvarchar](8) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DEPOSIT_TO] [nvarchar](10) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CHECKDATE] [datetime] NULL,
	[CARDDATE] [nvarchar](15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CHECKNUM] [nvarchar](15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CARDNUMBER] [nvarchar](16) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ID] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
