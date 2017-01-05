
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO

ALTER VIEW [dbo].[CMIS_Off_Hd_Deposited]
AS
SELECT     dbo.CMIS_Off_Hd.[ORDER], dbo.CMIS_Off_Hd.OR_NUM, dbo.CMIS_Off_Hd.OR_DATE, dbo.CMIS_Off_Hd.OR_AMT, dbo.CMIS_Off_Hd.DISCOUNT, 
                      dbo.CMIS_Off_Hd.TAX, dbo.CMIS_Off_Hd.CONSUMED, dbo.CMIS_Off_Hd.BAYADAMT, dbo.CMIS_Off_Hd.SUKLI, dbo.CMIS_Off_Hd.CASHAMOUNT, 
                      dbo.CMIS_Off_Hd.CHKAMOUNT, dbo.CMIS_Off_Hd.CARDAMOUNT, dbo.CMIS_Off_Hd.CUSCDE, dbo.CMIS_Off_Hd.CUSNAME, 
                      dbo.CMIS_Off_Hd.ADJUSTED, dbo.CMIS_Off_Hd.VARIOUS, dbo.CMIS_Off_Hd.PAIDNA, dbo.CMIS_Off_Hd.CANCEL, dbo.CMIS_Off_Hd.FORCE, 
                      dbo.CMIS_Off_Hd.DEPOSIT, dbo.CMIS_Off_Hd.BOUNCE, dbo.CMIS_Off_Hd.TOF, dbo.CMIS_Off_Hd.BANKCODE, dbo.CMIS_Off_Hd.BANKBRANCH, 
                      dbo.CMIS_Off_Hd.TSEKE, dbo.CMIS_Off_Hd.CARDBNKCDE, dbo.CMIS_Off_Hd.CARDNUMBER, dbo.CMIS_Off_Hd.TSEKLASE, 
                      dbo.CMIS_Off_Hd.CUTDATE, dbo.CMIS_Off_Hd.CHECKDATE, dbo.CMIS_Off_Hd.CARDDATE, dbo.CMIS_Off_Hd.DATECANCEL, 
                      dbo.CMIS_Off_Hd.TIMECANCEL, dbo.CMIS_Off_Hd.WHOCANCEL, dbo.CMIS_Off_Hd.DATECREATE, dbo.CMIS_Off_Hd.TIMECREATE, 
                      dbo.CMIS_Off_Hd.WHOCREATE, dbo.CMIS_Off_Hd.PRINTED, dbo.CMIS_Off_Hd.VALID, dbo.CMIS_Off_Hd.STATUS, dbo.CMIS_Off_Hd.VAT, 
                      dbo.CMIS_Off_Hd.ID, dbo.CMIS_BankDepo.DEPOSIT_TO, dbo.ALL_BANKS.BankName, dbo.ALL_BANKS.AcctCode AS BankAccountNo, 
                      dbo.CMIS_BankDepo.DATDEPOSIT
FROM         dbo.ALL_BANKS RIGHT OUTER JOIN
                      dbo.CMIS_BankDepo ON dbo.ALL_BANKS.BankCode = dbo.CMIS_BankDepo.DEPOSIT_TO LEFT OUTER JOIN
                      dbo.CMIS_Off_Hd ON dbo.CMIS_BankDepo.OR_NUM = dbo.CMIS_Off_Hd.OR_NUM
WHERE     (dbo.CMIS_Off_Hd.CARDNUMBER IS NULL)
GO

SET ANSI_NULLS OFF
GO
SET QUOTED_IDENTIFIER OFF
GO

