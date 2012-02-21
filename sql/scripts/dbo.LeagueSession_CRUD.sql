USE lakevillejuniors
GO

--region delete sproc
IF OBJECT_ID('[dbo].[up_DeleteLeagueSession]') IS NOT NULL
DROP PROC [dbo].[up_DeleteLeagueSession]
GO

CREATE PROC [dbo].[up_DeleteLeagueSession]
	@LeagueSessionID tinyint
AS 
/*************************************************
	dbo.up_DeleteLeagueSession:
	----------------------------------------------
	
*************************************************/
SET NOCOUNT ON

DELETE FROM dbo.LeagueSession
WHERE LeagueSession.LeagueSessionID = @LeagueSessionID
IF @@ERROR <> 0 RETURN(-1)

SET NOCOUNT OFF
RETURN(0)
GO

IF USER_ID('webapplication') IS NOT NULL
GRANT EXEC ON [dbo].[up_DeleteLeagueSession]
TO [webapplication]
GO
--endregion

--region select sproc
IF OBJECT_ID('[dbo].[up_GetLeagueSession]') IS NOT NULL
DROP PROC [dbo].[up_GetLeagueSession]
GO

CREATE PROC [dbo].[up_GetLeagueSession]
	@LeagueSessionID tinyint
AS 
/*************************************************
	dbo.up_GetLeagueSession:
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
GRANT EXEC ON [dbo].[up_GetLeagueSession]
TO [webapplication]
GO
--endregion

--region insert sproc
IF OBJECT_ID('[dbo].[up_InsertLeagueSession]') IS NOT NULL
DROP PROC [dbo].[up_InsertLeagueSession]
GO

CREATE PROC [dbo].[up_InsertLeagueSession]
	@Name varchar(200)
	,@Description varchar(2000) = NULL
	,@DisplayOrder tinyint
	,@Price money
	,@NewID tinyint OUTPUT
AS 
/*************************************************
	dbo.up_InsertLeagueSession:
	----------------------------------------------
	
*************************************************/
SET NOCOUNT ON

DECLARE @Err int

INSERT dbo.LeagueSession
	(Name
	,Description
	,DisplayOrder
	,Price
	)
VALUES
	(@Name
	,@Description
	,@DisplayOrder
	,@Price
	)
SELECT @NewID = SCOPE_IDENTITY(), @Err = @@ERROR
IF @Err <> 0 RETURN(-1)

SET NOCOUNT OFF
RETURN(0)
GO

IF USER_ID('webapplication') IS NOT NULL
GRANT EXEC ON [dbo].[up_InsertLeagueSession]
TO [webapplication]
GO

IF OBJECT_ID('[dbo].[up_UpdateLeagueSession]') IS NOT NULL
DROP PROC [dbo].[up_UpdateLeagueSession]
GO
--endregion

--region update sproc
CREATE PROC [dbo].[up_UpdateLeagueSession]
	@LeagueSessionID tinyint
	,@Name varchar(200)
	,@Description varchar(2000) = NULL
	,@DisplayOrder tinyint
	,@Price money
AS 
/*************************************************
	dbo.up_UpdateLeagueSession:
	----------------------------------------------
	
*************************************************/
SET NOCOUNT ON

UPDATE dbo.LeagueSession
SET
	Name = @Name
	,Description = @Description
	,DisplayOrder = @DisplayOrder
	,Price = @Price
WHERE LeagueSessionID = @LeagueSessionID
IF @@ERROR <> 0 RETURN(-1)

SET NOCOUNT OFF
RETURN(0)
GO

IF USER_ID('webapplication') IS NOT NULL
GRANT EXEC ON [dbo].[up_UpdateLeagueSession]
TO [webapplication]
GO
--endregion


