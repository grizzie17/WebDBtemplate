-- -----------------------------------------------------
-- Table [config]
-- -----------------------------------------------------

SELECT *
	INTO [config_old]
	FROM [config];

DROP TABLE [config];



CREATE TABLE [config] (
	[RID] COUNT,
	[Name] TEXT(64) NOT NULL UNIQUE,
	[Access] TEXT(4) NOT NULL DEFAULT 'SRVe',
	[Description] TEXT(128) NOT NULL,
	[Type] INTEGER NOT NULL DEFAULT 0,		-- 0=string, 1=integer
	[Maximum] INTEGER NULL DEFAULT NULL,	-- string length, max numeric value
	[DataValue] TEXT(255) NULL DEFAULT NULL
);

CREATE INDEX [configName] ON [config] ([Name]);

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('Domain', 'SRVe', 'Domain Name', 0, NULL, 'Unknown.tld' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[Domain]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'Domain';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteName', 'CFGe', 'Site Name', 0, NULL, 'Unknown Site' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[SiteName]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'SiteName';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteTitle', 'CFGe', 'Site Title (page header)', 0, NULL, 'Unknown Site' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[SiteTitle]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'SiteTitle';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('ShortSiteName', 'CFGe', 'Site Org. Designation', 0, NULL, 'Chapter XX-?' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[ShortSiteName]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'ShortSiteName';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteMotto', 'CFGe', 'Site Motto', 0, NULL, 'This is Your Motto!' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[SiteMotto]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'SiteMotto';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteCity', 'CFGe', 'Site City, State', 0, NULL, 'Anywhereville, XX' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[SiteCity]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'SiteCity';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('CopyrightStartYear', 'CFGe', 'Site Copyright Starting Year', 1, 9999, '2001' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[CopyrightStartYear]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'CopyrightStartYear';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('MailServer', 'SRVe', 'Mail Server', 0, NULL, 'mail.unknown.tld' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[MailServer]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'MailServer';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('MailboxUpcomingEvents', 'SRVe', 'Mailbox Upcoming Events', 0, NULL, 'UpcomingEvents@Unknown.tld' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[MailboxUpcomingEvents]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'MailboxUpcomingEvents';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('MailboxNewsletter', 'SRVe', 'Mailbox Newsletter Notification', 0, NULL, 'NewsletterNotice@Unknown.tld' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[MailboxNewsletter]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'MailboxNewsletter';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('AnnonUser', 'SRVe', 'Notification Email', 0, NULL, 'notifications@unknown.tld' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[MailUser]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'AnnonUser';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('AnnonPW', 'SRVe', 'Notification Password', 0, NULL, '' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[MailPW]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'AnnonPW';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteZip', 'CFGe', 'Site ZIP (Used for Weather)', 0, 5, '35801' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[SiteZip]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'SiteZip';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteChapter', 'CFGe', 'Site Chapter', 0, 5, 'AL-X' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[SiteChapter]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'SiteChapter';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteChapterID', 'CFGe', 'Site Chapter ID', 0, 2, 'X' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[SiteChapterID]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'SiteChapterID';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteTabNewsletters', 'CFGe', 'Tab Newsletters', 0, NULL, 'Newsletters' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[SiteTabNewsletters]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'SiteTabNewsletters';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteTabAnnouncements', 'CFGe', 'Tab Announcements', 0, NULL, 'Extra!' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[SiteTabAnnouncements]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'SiteTabAnnouncements';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteRSSAnnounceInclude', 'CFGe', 'RSS Announcements (Filter Include)', 0, NULL, '' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[SiteRSSAnnounceInclude]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'SiteRSSAnnounceInclude';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteRSSAnnounceExclude', 'CFGe', 'RSS Announcements (Filter Exclude)', 0, NULL, 'calendar,personal,plaque' );

UPDATE [config]
	SET [config].[DataValue] = [config_old].[SiteRSSAnnounceExclude]
	FROM [config_old], [config]
	WHERE [config].[Name] = 'SiteRSSAnnounceExclude';

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('UpcomingEventsSchedule', 'CFGe', 'Email Schedule for Upcoming Events', 0, 8, 'mo,fr' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('MetaKeywords', 'CFGe', 'Homepage Search Keywords', 0, NULL, 'gwrra, goldwing, gold wing road riders association, honda, motorcycle' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('CalendarCategories', 'CFGe', 'Homepage Calendar Categories', 0, NULL, '' );


DROP TABLE [config_old];



CREATE TABLE [images] (
  [RID] COUNT,
  [File] TEXT(80) NULL DEFAULT NULL,
  [Mime] TEXT(40) NULL DEFAULT NULL,
  [DateCreated] DATETIME NOT NULL,
  [Data] MEMO NULL DEFAULT NULL
);

/* Foreign keys for Images */
ALTER TABLE [images] ADD CONSTRAINT [PicturesImages] FOREIGN KEY (PictureID) REFERENCES [Pictures] (PictureID);

INSERT INTO [images]
	FROM [Pictures]
	SET [images].[RID] = [Pictures].[PictureID],
	SET [images].[File] = [Pictures].[File],
	SET [images].[DateCreated] = NOW(),




ALTER TABLE [Categories]
	ALTER COLUMN [SortName] TEXT(32);
ALTER TABLE [Categories]
	ALTER COLUMN [Icon] INTEGER NULL DEFAULT NULL;
ALTER TABLE [Categories]
	ADD COLUMN [Revision] INTEGER NOT NULL DEFAULT 0;




UPDATE [Version]
	SET Major = 8, Minor = 0, Fixes = 0;

