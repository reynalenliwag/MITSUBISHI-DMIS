SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO


ALTER    TRIGGER TrigUpdateSAEInfo on hrms_empinfo  for insert ,Update,DELETE
as
Declare @ISSAE BIT
Declare @EMPID INT
	Select @EMPID=ID, @ISSAE = is_sae from inserted 


		DELETE FROM smis_salesteam where saeid=@EMPID
		If @ISSAE=1 
		BEGIN			
			insert into smis_salesteam (Saeid)values(@EMPID)
		END 


GO
SET QUOTED_IDENTIFIER OFF 
GO
SET ANSI_NULLS ON 
GO

