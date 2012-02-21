use lakevillejuniors
go

IF OBJECT_ID('[dbo].[vw_OrphanedRegistrations]') IS NOT NULL
DROP VIEW [dbo].[vw_OrphanedRegistrations]
GO

CREATE VIEW [dbo].[vw_OrphanedRegistrations] AS

select
	NameFirstPlayer
	,NameLastPlayer
	,NameFirstParent1
	,NameLastParent1
	,Email
	,Grade
	,IsPaymentConfirmed
	,RegisPaid
	,PayPalPaymentStatus
from dbo.registrations r
where	IsOnlineRegistration = 1
and		IsPaymentConfirmed = 0

GO

/*
SELECT * FROM vw_OrphanedRegistrations ORDER BY NameLastPlayer, NameFirstPlayer
*/