USE [lakevillejuniors]
GO
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'webapplication' AND type = 'R')
CREATE ROLE [webapplication]
GO
IF NOT EXISTS (SELECT * FROM sys.database_principals WHERE name = N'weblogin')
CREATE USER [weblogin] FOR LOGIN [weblogin] WITH DEFAULT_SCHEMA=[dbo]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[Registrations]') AND type in (N'U'))
BEGIN
CREATE TABLE [dbo].[Registrations](
	[RegistrationID] [int] IDENTITY(1,1) NOT NULL,
	[NameFirstPlayer] [varchar](50) NOT NULL,
	[NameLastPlayer] [varchar](50) NOT NULL,
	[NameFirstParent1] [varchar](50) NOT NULL,
	[NameLastParent1] [varchar](50) NOT NULL,
	[AddressLine1] [varchar](100) NOT NULL,
	[AddressLine2] [varchar](100) NULL,
	[City] [varchar](50) NOT NULL,
	[StateID] [varchar](2) NULL,
	[Zip] [varchar](50) NOT NULL,
	[Phone] [varchar](50) NOT NULL,
	[Email] [varchar](100) NOT NULL,
	[School] [varchar](3) NOT NULL,
	[TShirtSize] [varchar](5) NOT NULL,
	[Grade] [tinyint] NOT NULL,
	[IsParentHelper] [tinyint] NOT NULL CONSTRAINT [DF_Registrations_IsParentHelper]  DEFAULT ((0)),
	[Notes] [varchar](2000) NULL,
	[DateCreated] [smalldatetime] NOT NULL CONSTRAINT [DF_Registrations_DateCreated]  DEFAULT (getdate()),
 CONSTRAINT [PK_Registrations] PRIMARY KEY CLUSTERED 
(
	[RegistrationID] ASC
)WITH (IGNORE_DUP_KEY = OFF) ON [PRIMARY]
) ON [PRIMARY]
END
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[f_PhoneClean]') AND type in (N'FN', N'IF', N'TF', N'FS', N'FT'))
BEGIN
execute dbo.sp_executesql @statement = N'CREATE FUNCTION [dbo].[f_PhoneClean](@Phone [varchar](14))
RETURNS [char](10) WITH EXECUTE AS CALLER
AS
BEGIN
	SET @Phone = REPLACE(@Phone, '' '', '''')
	SET @Phone = REPLACE(@Phone, ''('', '''')
	SET @Phone = REPLACE(@Phone, '')'', '''')
	SET @Phone = REPLACE(@Phone, ''-'', '''')
	SET @Phone = REPLACE(@Phone, ''.'', '''')
	RETURN(@Phone)
END

' 
END

GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
IF NOT EXISTS (SELECT * FROM sys.objects WHERE object_id = OBJECT_ID(N'[dbo].[up_InsertRegistration]') AND type in (N'P', N'PC'))
BEGIN
EXEC dbo.sp_executesql @statement = N'
CREATE PROC [dbo].[up_InsertRegistration]
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
	,@IsParentHelper tinyint
	,@Notes varchar(2000)
AS

/***************************************************************************
	up_InsertRegistration:
	---------------------------------------------
	RETURN(0) - success
	
	Version Control Info
	---------------------------------------------
	$Author: stephen $
	$Modtime: 3/04/05 9:24a $
	$Revision: 4 $
	$Date: 3/04/05 9:24a $
	Created Date: 2006-02-26
	Created By: Stephen Willingham
****************************************************************************/
SET NOCOUNT ON

-- don''t allow duplicate records
IF EXISTS 
	(	SELECT r.RegistrationID
		FROM dbo.Registrations r
		WHERE	UPPER(r.NameFirstPlayer) = UPPER(@NameFirstPlayer)
		AND	UPPER(r.NameLastPlayer) = UPPER(@NameLastPlayer)
		AND	UPPER(r.Email) = UPPER(@Email)
		)
		RETURN(-1)


INSERT INTO [lakevillejuniors].[dbo].[Registrations]
	(	[NameFirstPlayer]
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
		)
VALUES
	(	@NameFirstPlayer
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
		,@IsParentHelper
		,@Notes
		,GETDATE()
		)
IF @@ERROR <> 0 RETURN(-2)

RETURN(0)
SET NOCOUNT OFF
' 
END
GO
GRANT EXECUTE ON [dbo].[up_InsertRegistration] TO [webapplication]
