USE [lakevillejuniors]
GO

IF OBJECT_ID('[dbo].[vw_HeadCoachCoverSheet]') IS NOT NULL
DROP VIEW [dbo].[vw_HeadCoachCoverSheet]
GO

CREATE VIEW [dbo].[vw_HeadCoachCoverSheet] AS
/***************************************************************************
	vw_HeadCoachCoverSheet
	------------------------------------------------------------------------
	Created:2007-09-09
	RETURN(0) - success

	Description:
****************************************************************************/

SELECT TOP(500)
	* 
FROM dbo.vw_MailingInfo
WHERE IsHeadCoach = 1
ORDER BY Team
GO

/*
SELECT * FROM dbo.vw_HeadCoachCoverSheet
*/
