USE lakevillejuniors
GO

CREATE PROC [dbo].[up_GetLeagueSessionList]
	@LeagueSessionID tinyint
AS 
/*************************************************
	dbo.up_GetLeagueSessionList:
	----------------------------------------------
	
*************************************************/
SET NOCOUNT ON

SELECT 
	LeagueSession.LeagueSessionID
	,LeagueSession.Name
	,LeagueSession.Description
	,LeagueSession.DisplayOrder
	,LeagueSession.Price
FROM dbo.LeagueSession
WHERE LeagueSession.LeagueSessionID = @LeagueSessionID

SET NOCOUNT OFF
RETURN(0)
GO

IF USER_ID('webapplication') IS NOT NULL
GRANT EXEC ON [dbo].[up_GetLeagueSessionList]
TO [webapplication]
GO
