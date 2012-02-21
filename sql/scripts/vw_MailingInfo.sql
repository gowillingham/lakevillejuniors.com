USE [lakevillejuniors]
GO

IF OBJECT_ID('[dbo].[vw_MailingInfo]') IS NOT NULL
DROP VIEW [dbo].[vw_MailingInfo]
GO

CREATE VIEW [dbo].[vw_MailingInfo] AS
/***************************************************************************
	vw_MailingInfo
	------------------------------------------------------------------------
	Created:2007-09-09
	RETURN(0) - success

	Description:
****************************************************************************/

SELECT TOP(500) 
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
	,ls.Price AS RegisFee
	,[RegisPaid]
	,[CoachTShirtSize]
	,[Team]
	,IsHeadCoach
	,ls.Price - RegisPaid AS Balance
	,dbo.f_PhoneFormat(Phone, 0) AS PhoneFormatted
	,CASE WHEN IsParentHelper = 1 THEN 'Yes!' ELSE 'No Thanks!' END AS IsParentHelperText
	,CASE WHEN ls.Price - RegisPaid = 0 THEN ''
	 ELSE	
		'You registered online, or we have not yet received your $70.00 ' +
		'league registration fee. Please bring payment to the first session this Saturday ' +  
		'(checks should be made out to Lakeville Juniors). '
	 END AS BalanceText
	,CASE WHEN IsHeadCoach = 1 THEN
		'You have been selected to be the lead parent helper for your daughter''s team! ' +
		'Watch the mail for your team roster and instructions for contacting your team ' +
		'(we''ll have you follow up with your team before the first session about times, locations, etc). '
	 ELSE '' END AS HeadCoachText
	,CASE WHEN COALESCE(IsHeadCoach, 0) = 0 AND IsParentHelper = 1 THEN
		'You have been selected to be an assistant coach for your team. ' +
		'Thanks for your willingness to help! '
	 ELSE '' END AS AssistantCoachText
	,CASE WHEN Session = 1 THEN 'Beginner (grade 1/2/3)' ELSE 'Intermediate (grade 3/4/5)' END AS SessionName
	,CASE WHEN DateCreated >= '2011-09-25' THEN 
		' Your registration was received after the registration deadline. ' +
		'It is possible that you will not receive your team shirt until the second session. '
	 ELSE '' END AS LateRegistrationText
	,Session
FROM [dbo].[Registrations] r
JOIN dbo.LeagueSession ls ON r.Session = ls.LeagueSessionID
WHERE (Session = 1) OR (Session = 2)
ORDER BY Email, DateCreated

GO

/*
SELECT * FROM dbo.vw_MailingInfo
*/
