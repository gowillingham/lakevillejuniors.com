USE [lakevillejuniors]
GO

IF OBJECT_ID('[dbo].[up_playerGetRegistrationListByEmail]') IS NOT NULL
DROP PROC [dbo].[up_playerGetRegistrationListByEmail]
GO

CREATE PROC [dbo].[up_playerGetRegistrationListByEmail]
	-- parameter list
	@Email varchar(100)
/***************************************************************************
	up_playerGetRegistrationListByEmail
	------------------------------------------------------------------------
	Created: 2007-11-29
	RETURN(0) - success

	Description:
****************************************************************************/
AS
SET NOCOUNT ON
/*
' 0-RegistrationID 1-RegistrationNumber 2-NameFirstPlayer 3-NameLastPlayer 4-NameFirstParent
' 5-NameLastParent 6-AddressLine1 7-AddressLine2 8-City 9-StateID 10-Zip
' 11-Phone 12-Email 13-school 14-TShirtSize 15-Grade 16-IsParentHelper 17-Notes 18-dateCreated
' 19-DateModified 20-IsOnlineRegistration 21-RegisFee 22-RegisPaid 23-CoachTShirtSize
' 24-Team 25-IsHeadCoach 26-IsPaymentConfirmed 27-PayPalTransactionId 28-PayPalIsSandbox
' 29-PayPalPaymentStatus 30-PayPalPaymentStatusReason 31-SessionName 32-SessionDescription
*/
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
	,ls.[Name]
	,ls.[Description]
FROM [dbo].[Registrations] r
JOIN [dbo].LeagueSession ls ON r.Session = ls.LeagueSessionID
WHERE r.Email = @Email
ORDER BY DateCreated

SET NOCOUNT OFF
RETURN(0)
GO

IF USER_ID('weblogin') IS NOT NULL
GRANT EXEC ON [dbo].[up_playerGetRegistrationListByEmail]
TO [weblogin]
GO

/*
DECLARE @retval INT
EXEC @retval = [dbo].[up_playerGetRegistrationListByEmail]
*/
