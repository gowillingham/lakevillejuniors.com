USE [lakevillejuniors]
GO

IF OBJECT_ID('[dbo].[up_playerGetRegistration]') IS NOT NULL
DROP PROC [dbo].[up_playerGetRegistration]
GO

CREATE PROC [dbo].[up_playerGetRegistration]
	-- parameter list
	@RegistrationID uniqueidentifier
/***************************************************************************
	up_playerGetRegistration
	------------------------------------------------------------------------
	Created: 2007-11-29
	RETURN(0) - success

	Description:
****************************************************************************/
AS
SET NOCOUNT ON

SELECT 
	[RegistrationID]
	,RegistrationNumber
	,[NameFirstPlayer]
	,[NameLastPlayer]
	,[NameFirstParent1]
	,[NameLastParent1]
	,[AddressLine1]
	,[AddressLine2]
	,[City]
	,[StateID]
	,[Zip]
	,dbo.f_PhoneFormat([Phone], 1) AS Phone
	,[Phone] AS PhoneRaw
	,[Email]
	,[School]
	,[TShirtSize]
	,[Grade]
	,[IsParentHelper]
	,[Notes]
	,[DateCreated]
	,[DateModified]
	,[IsOnlineRegistration]
	,ls.Price AS RegisFee
	,[RegisPaid]
	,[CoachTShirtSize]
	,[Team]
	,[IsHeadCoach]
	,[IsPaymentConfirmed]
	,[PayPalTransactionID]
	,[PayPalIsSandbox]
	,PayPalPaymentStatus
	,PayPalPaymentStatusReason
	,Session
	,HasRelease
FROM [dbo].[Registrations] r
JOIN [dbo].LeagueSession ls ON r.Session = ls.LeagueSessionID
WHERE RegistrationID = @RegistrationID

SET NOCOUNT OFF
RETURN(0)
GO

IF USER_ID('weblogin') IS NOT NULL
GRANT EXEC ON [dbo].[up_playerGetRegistration]
TO [weblogin]
GO

/*
DECLARE @retval INT
EXEC @retval = [dbo].[up_playerGetRegistration]
*/
