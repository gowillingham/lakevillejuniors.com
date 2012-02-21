USE [lakevillejuniors]
GO

IF OBJECT_ID('[dbo].[up_playerInsertRegistration]') IS NOT NULL
DROP PROC [dbo].[up_playerInsertRegistration]
GO

CREATE PROC [dbo].[up_playerInsertRegistration]
	-- parameter list
	@NameFirstPlayer varchar(50)
	,@NameLastPlayer varchar(50)
	,@NameFirstParent1 varchar(50)
	,@NameLastParent1 varchar(50)
	,@AddressLine1 varchar(100)
	,@AddressLine2 varchar(100)
	,@City varchar(50)
	,@StateID varchar(2)
	,@Zip varchar(50)
	,@Phone varchar(50)
	,@Email varchar(100)
	,@School varchar(3)
	,@TShirtSize varchar(5)
	,@Grade tinyint
	,@Session tinyint = NULL
	,@IsParentHelper tinyint
	,@Notes varchar(2000)
	,@DateCreated smalldatetime
	,@IsOnlineRegistration tinyint
	,@RegisPaid money = NULL
	,@CoachTShirtSize varchar(5) = NULL
	,@Team varchar(25) = NULL
	,@IsHeadCoach tinyint = NULL
	,@IsPaymentConfirmed tinyint 
	,@PayPalTransactionID varchar(256) = NULL
	,@PayPalIsSandbox tinyint = NULL
	,@PayPalPaymentStatus varchar(256) = NULL
	,@PayPalPaymentStatusReason varchar(256) = NULL
	,@HasRelease tinyint = NULL
	,@NewID uniqueidentifier OUTPUT
/***************************************************************************
	up_playerInsertRegistration
	------------------------------------------------------------------------
	Created: 2007-11-29
	RETURN(0) - success

	Description:
****************************************************************************/
AS
SET NOCOUNT ON

SET @NewID = NEWID()

INSERT INTO [dbo].[Registrations]
	([RegistrationID]
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
	,[Session]
	,[IsParentHelper]
	,[Notes]
	,[DateCreated]
	,[DateModified]
	,[IsOnlineRegistration]
	,[RegisPaid]
	,[CoachTShirtSize]
	,[Team]
	,[IsHeadCoach]
	,[IsPaymentConfirmed]
	,[PayPalTransactionID]
	,[PayPalIsSandbox]
	,PayPalPaymentStatus
	,PayPalPaymentStatusReason
	,HasRelease
	)
VALUES
	(@NewID
	,@NameFirstPlayer
	,@NameLastPlayer
	,@NameFirstParent1
	,@NameLastParent1
	,@AddressLine1
	,@AddressLine2
	,@City
	,@StateID
	,@Zip
	,dbo.f_PhoneClean(@Phone)
	,@Email
	,@School
	,@TShirtSize
	,@Grade
	,@Session
	,@IsParentHelper
	,@Notes
	,@DateCreated
	,@DateCreated
	,@IsOnlineRegistration
	,@RegisPaid
	,@CoachTShirtSize
	,@Team
	,@IsHeadCoach
	,@IsPaymentConfirmed
	,@PayPalTransactionID
	,@PayPalIsSandbox
	,@PayPalPaymentStatus
	,@PayPalPaymentStatusReason
	,@HasRelease
	)

SET NOCOUNT OFF
RETURN(0)
GO

IF USER_ID('weblogin') IS NOT NULL
GRANT EXEC ON [dbo].[up_playerInsertRegistration]
TO [weblogin]
GO

/*
DECLARE @retval INT
EXEC @retval = [dbo].[up_playerInsertRegistration]
*/
