USE [lakevillejuniors]
GO

IF OBJECT_ID('[dbo].[up_GetPlayerMailingInfo]') IS NOT NULL
DROP PROC [dbo].[up_GetPlayerMailingInfo]
GO

CREATE PROC [dbo].[up_GetPlayerMailingInfo]
	-- parameter list
	
/***************************************************************************
	up_GetPlayerMailingInfo
	------------------------------------------------------------------------
	Created:2007-09-09
	RETURN(0) - success

	Description:
****************************************************************************/
AS
SET NOCOUNT ON

SELECT 
	[RegistrationID]
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
	,[RegisFee]
	,[RegisPaid]
	,[CoachTShirtSize]
	,[Team]
	,IsHeadCoach
	,RegisFee - RegisPaid AS Balance
	,dbo.f_PhoneFormat(Phone, 0) AS PhoneFormatted
	,CASE WHEN IsParentHelper = 1 THEN 'Yes!' ELSE 'No Thanks!' END AS IsParentHelperText
	,CASE WHEN RegisFee - RegisPaid = 0 THEN ''
	 ELSE	
		'You registered online, or we have not yet received your $60.00 ' +
		'league registration fee. Please bring payment to the first session this Saturday ' +  
		'(checks should be made out to Lakeville Juniors).'
	 END AS BalanceText
	,CASE WHEN IsHeadCoach = 1 THEN
		'You have been selected to be the head coach for your daughter''s team! ' +
		'You will find the head coach instructions enclosed with the league confirmation mailing ' +
		'(head coach instructions are not included with your email confirmation).'
	 ELSE '' END AS HeadCoachText
	,CASE WHEN IsHeadCoach = 0 AND IsParentHelper = 1 THEN
		'You have been selected to be an assistant coach for your daughter''s team. ' +
		'Thanks for your willingness to help!'
	 ELSE '' END AS AssistantCoachText
	,CASE WHEN Grade < 4 THEN 'Session I' ELSE 'Session II' END AS Session
FROM [dbo].[Registrations] r


SET NOCOUNT OFF
RETURN(0)
GO

IF USER_ID('webapplication') IS NOT NULL
GRANT EXEC ON [dbo].[up_GetPlayerMailingInfo]
TO [webapplication]
GO

DECLARE @retval INT
EXEC @retval = [dbo].[up_GetPlayerMailingInfo]
/*
*/



