USE [lakevillejuniors]
GO

IF OBJECT_ID('[dbo].[up_GetRosters]') IS NOT NULL
DROP PROC [dbo].[up_GetRosters]
GO

CREATE PROC [dbo].[up_GetRosters]
	-- parameter list
	@TeamColor varchar(25)
	,@SessionNumber tinyint
/***************************************************************************
	up_GetRosters
	------------------------------------------------------------------------
	Created:2007-09-10
	RETURN(0) - success

	Description:
****************************************************************************/
AS
SET NOCOUNT ON

Select 
	NameLastPlayer + ', ' + NameFirstPlayer AS Player
	,NameLastParent1 + ', ' + NameFirstParent1 AS Parent
	,dbo.f_PhoneFormat(Phone, 1) AS Phone
	,Team
	,CASE WHEN IsParentHelper = 1 THEN 'Yes' ELSE '' END AS [Parent Volunteer]
	,SessionName
	,Grade
	,School
FROM dbo.vw_MailingInfo
WHERE	Team = @TeamColor 
AND		Session = @SessionNumber
ORDER BY NameLastPlayer, NameFirstPlayer

SET NOCOUNT OFF
RETURN(0)
GO

IF USER_ID('webapplication') IS NOT NULL
GRANT EXEC ON [dbo].[up_GetRosters]
TO [webapplication]
GO

DECLARE @retval INT
EXEC @retval = [dbo].[up_GetRosters] 'Brown', 1
/*
*/



