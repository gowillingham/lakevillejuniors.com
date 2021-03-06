USE lakevillejuniors
GO

/*
DELETE registrations
*/

SELECT 
	[RegistrationID]
	,[RegistrationNumber]
	,[IsPaymentConfirmed]
	,[RegisFee]
	,[RegisPaid]
	,[PayPalTransactionID]
	,[PayPalIsSandbox]
	,[NameFirstPlayer]
	,[NameLastPlayer]
	,[NameFirstParent1]
	,[NameLastParent1]
	,[AddressLine1]
	,[AddressLine2]
	,[City]
	,[StateID]
	,[Zip]
	,[Phone]
	,[Email]
	,[School]
	,[TShirtSize]
	,[Grade]
	,[IsParentHelper]
	,[Notes]
	,[DateCreated]
	,[IsOnlineRegistration]
	,[CoachTShirtSize]
	,[Team]
	,[IsHeadCoach]
	,[DateModified]
FROM [lakevillejuniors].[dbo].[Registrations]

