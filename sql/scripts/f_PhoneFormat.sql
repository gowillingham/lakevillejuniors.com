USE [lakevillejuniors]
GO

IF OBJECT_ID ('[dbo].[f_PhoneFormat]') IS NOT NULL
   DROP FUNCTION [dbo].[f_PhoneFormat]
GO

CREATE FUNCTION [dbo].[f_PhoneFormat]
	(@OldPhone char(10), @Style tinyint)
RETURNS VARCHAR(14)
WITH EXECUTE AS CALLER
/***************************************************************************
	f_PhoneFormat: Accept 10-digit phone and format according to style selected.
	---------------------------------------------
	RETURNS: VARCHAR(14)

	@Style = 1: (XXX) XXX-XXXX
	@Style = 2: XXX-XXX-XXXX
	@Style = 3: XXX.XXX.XXXX
	@Style = 4: XXX XXX XXXX
	Else Do Nothing

	Version Control Info
	---------------------------------------------
	$Author: stephen $
	$Modtime: 11/18/05 3:10p $
	$Revision: 2 $
	$Date: 11/18/05 3:10p $
	Created Date: 2004-12-03
	Created By: Stephen Willingham
****************************************************************************/
AS BEGIN

	DECLARE @NewPhone VARCHAR(14)
	IF LEN(@OldPhone) = 0 RETURN(@NewPhone)
	IF @OldPhone IS NULL RETURN(@NewPhone)
	SET @NewPhone = @OldPhone
	IF @Style = 1 BEGIN
		SET @NewPhone = STUFF(@NewPhone, 1 , 0 , '(')
		SET @NewPhone = STUFF(@NewPhone, 5 , 0 , ') ')
		SET @NewPhone = STUFF(@NewPhone, 10, 0, '-')
	END
	ELSE IF @Style = 2 BEGIN
		SET @NewPhone = STUFF(@NewPhone, 4, 0, '-')
		SET @NewPhone = STUFF(@NewPhone, 8, 0, '-')
	END
	ELSE IF @Style = 3 BEGIN
		SET @NewPhone = STUFF(@NewPhone, 4, 0, '.')
		SET @NewPhone = STUFF(@NewPhone, 8, 0, '.')
	END
	ELSE IF @Style = 4 BEGIN
		SET @NewPhone = STUFF(@NewPhone, 4, 0, ' ')
		SET @NewPhone = STUFF(@NewPhone, 8, 0, ' ')
	END
	ELSE SET @NewPhone = @OldPhone
	
	RETURN(@NewPhone)
END
GO
