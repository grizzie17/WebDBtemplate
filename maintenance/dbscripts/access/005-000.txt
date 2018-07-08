# alter the Config table


ALTER TABLE [Config]
	ADD COLUMN [SiteRSSAnnounceInclude] TEXT(64)  WITH COMPRESSION NULL;

ALTER TABLE [Config]
	ADD COLUMN [SiteRSSAnnounceExclude] TEXT(64)  WITH COMPRESSION NULL;

UPDATE [Config]
	SET [SiteRSSAnnounceExclude] = "calendar,personal,plaque";



UPDATE Version
	SET Major = 5, Minor = 1, Fixes = 0;

