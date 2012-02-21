USE [lakevillejuniors]
GO

IF OBJECT_ID('[dbo].[up_playerUpdateRegistration]') IS NOT NULL
DROP PROC [dbo].[up_playerUpdateRegistration]
GO

CREATE PROC [dbo].[up_playerUpdateRegistration]
	-- parameter list
	@RegistrationID uniqueidentifier
	,@NameFirstPlayer varchar(50)
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
	,@DateModified smalldatetime
	,@IsOnlineRegistration tinyint
	,@RegisPaid money
	,@CoachTShirtSize varchar(5)
	,@Team varchar(25)
	,@IsHeadCoach tinyint
	,@IsPaymentConfirmed tinyint
	,@PayPalTransactionID varchar(256) = NULL
	,@PayPalIsSandbox tinyint = NULL  
	,@PayPalPaymentStatus varchar(256) = NULL
	,@PayPalPaymentStatusReason varchar(256) = NULL
	,@HasRelease tinyint = NULL
/***************************************************************************
	up_playerUpdateRegistration
	------------------------------------------------------------------------
	Created: 2007-11-19
	RETURN(0) - success

	Description:
****************************************************************************/
AS
SET NOCOUNT ON

UPDATE [dbo].[Registrations]
SET [NameFirstPlayer] = @NameFirstPlayer
	,[NameLastPlayer] = @NameLastPlayer
	,[NameFirstParent1] = @NameFirstParent1
	,[NameLastParent1] = @NameLastParent1
	,[AddressLine1] = @AddressLine1
	,[AddressLine2] = @AddressLine2
	,[City] = @City
	,[StateID] = @StateID
	,[Zip] = @Zip
	,[Phone] = dbo.f_PhoneClean(@Phone)
	,[Email] = @Email
	,[School] = @School
	,[TShirtSize] = @TShirtSize
	,[Grade] = @Grade
	,[Session] = @Session
	,[IsParentHelper] = @IsParentHelper
	,[Notes] = @Notes
	,[IsOnlineRegistration] = @IsOnlineRegistration
	,[RegisPaid] = @RegisPaid
	,[CoachTShirtSize] = @CoachTShirtSize
	,[Team] = @Team
	,[IsHeadCoach] = @IsHeadCoach
	,[IsPaymentConfirmed] = @IsPaymentConfirmed
	,[DateModified] = @DateModified
	,[PayPalTransactionID] = @PayPalTransactionID
	,[PayPalIsSandbox] = @PayPalIsSandbox
	,PayPalPaymentStatus = @PayPalPaymentStatus
	,[PayPalPaymentStatusReason] = @PayPalPaymentStatusReason
	,[HasRelease] = @HasRelease
WHERE RegistrationID = @RegistrationID

SET NOCOUNT OFF
RETURN(0)
GO

IF USER_ID('weblogin') IS NOT NULL
GRANT EXEC ON [dbo].[up_playerUpdateRegistration]
TO [weblogin]
GO

/*
DECLARE @retval INT
EXEC @retval = [dbo].[up_playerUpdateRegistration]
*/
