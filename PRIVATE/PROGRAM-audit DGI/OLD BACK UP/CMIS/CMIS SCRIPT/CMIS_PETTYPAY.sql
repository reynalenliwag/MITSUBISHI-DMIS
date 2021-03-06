USE [DMIS]
GO
/****** Object:  Table [dbo].[CMIS_PETTYPAY]    Script Date: 01/29/2010 14:35:50 ******/
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[CMIS_PETTYPAY](
	[TYPE] [nvarchar](1) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CUTDATE] [datetime] NULL,
	[INCASHDATE] [datetime] NULL,
	[BANKCODE] [nvarchar](3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CHKNUMBER] [nvarchar](15) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[CHKDATE] [datetime] NULL,
	[CHKAMOUNT] [decimal](18, 2) NULL,
	[PAYMENTAMT] [decimal](18, 2) NULL,
	[INCASHAMT] [decimal](18, 2) NULL,
	[TIMEINCASH] [nvarchar](12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DATECREATE] [datetime] NULL,
	[TIMECREATE] [nvarchar](12) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[DEPOSIT] [bit] NULL,
	[TSEKLASE] [nvarchar](3) COLLATE SQL_Latin1_General_CP1_CI_AS NULL,
	[ID] [int] IDENTITY(1,1) NOT NULL
) ON [PRIMARY]
