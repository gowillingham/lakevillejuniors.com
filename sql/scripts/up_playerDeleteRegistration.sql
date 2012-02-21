USE [lakevillejuniors]
GO

IF OBJECT_ID('[dbo].[up_playerDeleteRegistration]') IS NOT NULL
DROP PROC [dbo].[up_playerDeleteRegistration]
GO

CREATE PROC [dbo].[up_playerDeleteRegistration]
	-- parameter list
	@RegistrationID uniqueidentifier
/***************************************************************************
	up_playerDeleteRegistration
	------------------------------------------------------------------------
	Created: 2007-11-29
	RETURN(0) - success

	Description:
****************************************************************************/
AS
SET NOCOUNT ON

DELETE dbo.Registrations
WHERE RegistrationID = @RegistrationID

SET NOCOUNT OFF
RETURN(0)
GO

IF USER_ID('weblogin') IS NOT NULL
GRANT EXEC ON [dbo].[up_playerDeleteRegistration]
TO [weblogin]
GO

/*
DECLARE @retval INT
EXEC @retval = [dbo].[up_playerDeleteRegistration]
*/
