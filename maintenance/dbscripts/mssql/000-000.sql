
-- -----------------------------------------------------
-- Drop old tables
-- -----------------------------------------------------

DROP TABLE [authorizeresourcemap] ;
DROP TABLE [authorizeresource] ;

DROP TABLE [authorizelog] ;
DROP TABLE [authorizeremember] ;
DROP TABLE [authorizegroupmap] ;
DROP TABLE [authorizegroup];
DROP TABLE [authorizeusers] ;


DROP TABLE [config] ;
DROP TABLE [keywordpagemap];
DROP TABLE [keywords] ;

DROP TABLE [pagelistmap] ;
DROP TABLE [pagelistoptions] ;
DROP TABLE [pictures] ;
DROP TABLE [schedules] ;
DROP TABLE [tabcategorymap] ;
DROP TABLE [tabs] ;
DROP TABLE [pages] ;
DROP TABLE [categories] ;
DROP TABLE [version] ;
DROP TABLE [images] ;

DROP FUNCTION [LEAST];

CREATE FUNCTION dbo.LEAST( @d1 DATETIME, @d2 DATETIME ) RETURNS DATETIME
AS
BEGIN
	DECLARE @r DATETIME
	SELECT @r = CASE WHEN @d1 < @d2 THEN @d1 ELSE @d2 END
	RETURN @r
END;


-- -----------------------------------------------------
-- Table [authorizeusers]
-- -----------------------------------------------------

CREATE TABLE [authorizeusers]
(
	[RID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
	[Username] VARCHAR(32) NOT NULL UNIQUE,
	[Password] VARCHAR(32) NOT NULL,
	[Email] VARCHAR(128) NOT NULL UNIQUE,
	[Description] VARCHAR(255) NULL DEFAULT NULL,
	[DateCreated] DATETIME NOT NULL,
	[DateModified] DATETIME NOT NULL,
	[Disabled] BIT NOT NULL DEFAULT '0',
	[Lastname] VARCHAR(32) NULL DEFAULT NULL,
	[Firstname] VARCHAR(32) NULL DEFAULT NULL,
);


CREATE UNIQUE INDEX [authorizeusersRID] ON authorizeusers ([RID]);
CREATE UNIQUE INDEX [authorizeusersUsername] ON authorizeusers ([Username]);

INSERT INTO [authorizeusers]
	( [UserName], [Password], [Email], [DateCreated], [DateModified], [Lastname], [Firstname] )
	VALUES
	( 'grizzie17', 'grizzly17', 'john@grizzlyweb.com', GETDATE(), GETDATE(), 'Griswold', 'John' );

INSERT INTO [authorizeusers]
	( [UserName], [Password], [Email], [DateCreated], [DateModified], [Lastname], [Firstname] )
	VALUES
	( 'webby17', 'grizzly17', 'john@griwold.me', GETDATE(), GETDATE(), 'Griswold', 'John' );

INSERT INTO [authorizeusers]
	( [UserName], [Password], [Email], [DateCreated], [DateModified], [Lastname], [Firstname] )
	VALUES
	( 'admin123', '123.admin!321', 'sample@sample.net', GETDATE(), GETDATE(), 'Admin', 'Admin' );

-- -----------------------------------------------------
-- Table [authorizegroup]
-- -----------------------------------------------------


CREATE TABLE [authorizegroup] 
(
	[RID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
	[Name] VARCHAR(40) NOT NULL,
	[Description] VARCHAR(255) NULL DEFAULT NULL,
	[DateModified] DATETIME NOT NULL DEFAULT GETDATE(),
	[Disabled] BIT NOT NULL DEFAULT '0'
);

CREATE UNIQUE INDEX [authorizegroupName] ON authorizegroup ([Name]);
CREATE INDEX [authorizegroupRID] ON authorizegroup ([RID]);


-- -----------------------------------------------------
-- Table [authorizegroupmap]
-- -----------------------------------------------------

CREATE TABLE [authorizegroupmap]
(
  [RID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
  [UserID] INT NOT NULL,
  [GroupID] INT NOT NULL,
  CONSTRAINT [Usersauthorizegroupmap]
    FOREIGN KEY ([UserID])
    REFERENCES [authorizeusers] ([RID])
	ON DELETE CASCADE 
	ON UPDATE CASCADE,
  CONSTRAINT [Groupsauthorizegroupmap]
    FOREIGN KEY ([GroupID])
    REFERENCES [authorizegroup] ([RID])
	ON DELETE CASCADE 
	ON UPDATE CASCADE
);


CREATE INDEX [authorizegroupmapUserID] ON authorizegroupmap ([UserID]);
CREATE INDEX [authorizegroupmapGroupID] ON authorizegroupmap ([GroupID]);




-- -----------------------------------------------------
-- Table [authorizelog]
-- -----------------------------------------------------

CREATE TABLE [authorizelog] 
(
  [UserID] INT NOT NULL,
  [DateLogon] DATETIME NOT NULL,
  [DateFail] DATETIME NULL DEFAULT NULL,
  [FailCount] INT NOT NULL,
  CONSTRAINT [authorizeusersAuthorizeLog]
    FOREIGN KEY ([UserID])
    REFERENCES [authorizeusers] ([RID])
);

CREATE INDEX [authorizelogUserID] ON authorizelog ([UserID]);


-- -----------------------------------------------------
-- Table [authorizeremember]
-- -----------------------------------------------------

CREATE TABLE [authorizeremember] (
  [RID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
  [UserID] INT NOT NULL,
  [Cookie] VARCHAR(40) NOT NULL,
  [DateModified] DATETIME NOT NULL,
  CONSTRAINT [UsersAuthorizeRemember]
    FOREIGN KEY ([UserID])
    REFERENCES [authorizeusers] ([RID])
	ON DELETE CASCADE ON UPDATE CASCADE
);


CREATE INDEX [AuthorizeRememberRID] ON authorizeremember ([RID] ASC);
CREATE INDEX [AuthorizeRememberUserID] ON authorizeremember ([UserID] ASC);

-- -----------------------------------------------------
-- Table [authorizeresource]
-- -----------------------------------------------------

CREATE TABLE [authorizeresource] (
  [RID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
  [Name] VARCHAR(40) NOT NULL UNIQUE,
  [Abbrev] VARCHAR(4) NOT NULL UNIQUE,
  [Description] VARCHAR(255) NULL DEFAULT NULL,
  [DateModified] DATETIME NOT NULL,
  [Disabled] BIT NOT NULL DEFAULT '0'
);

CREATE INDEX [authorizeresourceName] ON [authorizeresource] ([Name] ASC);
CREATE INDEX [authorizeresourceRID] ON [authorizeresource] ([RID] ASC);

-- -----------------------------------------------------
-- Table [authorizeresourcemap]
-- -----------------------------------------------------

CREATE TABLE [authorizeresourcemap] (
  [RID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
  [GroupID] INT NOT NULL,
  [ResourceID] INT NOT NULL,
  CONSTRAINT [Resourcesauthorizeresourcemap]
    FOREIGN KEY ([ResourceID])
    REFERENCES [authorizeresource] ([RID])
	ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT [Groupsauthorizeresourcemap]
    FOREIGN KEY ([GroupID])
    REFERENCES [authorizegroup] ([RID])
	ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE INDEX [authorizeresourcemapRID] ON [authorizeresourcemap] ([RID] ASC);
CREATE INDEX [authorizeresourcemapGroupID] ON [authorizeresourcemap] ([GroupID] ASC);
CREATE INDEX [authorizeresourcemapResourceID] ON [authorizeresourcemap] ([ResourceID] ASC);

-- -----------------------------------------------------
-- Table [categories]
-- -----------------------------------------------------

CREATE TABLE [categories] (
  [CategoryID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
  [Name] VARCHAR(50) NOT NULL UNIQUE,
  [ShortName] VARCHAR(30) NULL DEFAULT NULL,
  [ShortDesc] VARCHAR(255) NULL DEFAULT NULL,
  [Icon] INT NULL DEFAULT NULL,	-- imageID
  [SortName] VARCHAR(32) NULL DEFAULT NULL,
  [Revision] INT NOT NULL DEFAULT 0,
  [Disabled] BIT NOT NULL DEFAULT '0'
);

INSERT INTO [categories] ( [Name], [ShortDesc] ) VALUES ( '@Library', 'Images for global use' );
INSERT INTO [categories] ( [Name] ) VALUES ( '@External' );
INSERT INTO [categories] ( [Name], [ShortDesc] ) VALUES ( 'Home', 'Special content on home page' );
INSERT INTO [categories] ( [Name], [ShortDesc] ) VALUES ( 'Calendar', 'Special content on calendar page' );
INSERT INTO [categories] ( [Name], [ShortDesc] ) VALUES ( 'Announcements', 'Content for Home and Extra! pages' );
INSERT INTO [categories] ( [Name], [ShortDesc] ) VALUES ( 'Newsletters', 'Content for Newsletters pages' );


-- -----------------------------------------------------
-- Table [config]
-- -----------------------------------------------------


CREATE TABLE [config] (
	[RID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
	[Name] VARCHAR(64) NOT NULL UNIQUE,
	[Access] VARCHAR(4) NOT NULL DEFAULT 'SRVe',
	[Description] VARCHAR(128) NOT NULL,
	[Type] INT NOT NULL DEFAULT 0,		-- 0=string, 1=integer
	[Maximum] INT NULL DEFAULT NULL,	-- string length, max numeric value
	[DataValue] VARCHAR(255) NULL DEFAULT NULL
);

CREATE INDEX [configName] ON [config] ([Name]);

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('Domain', 'SRVe', 'Domain Name', 0, NULL, 'Unknown.tld' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteName', 'CFGe', 'Site Name', 0, NULL, 'Unknown Site' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteTitle', 'CFGe', 'Site Title (page header)', 0, NULL, 'Unknown Site' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('ShortSiteName', 'CFGe', 'Site Org. Designation', 0, NULL, 'Chapter XX-?' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteMotto', 'CFGe', 'Site Motto', 0, NULL, 'This is Your Motto!' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteCity', 'CFGe', 'Site City, State', 0, NULL, 'Anywhereville, XX' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('CopyrightStartYear', 'CFGe', 'Site Copyright Starting Year', 1, 9999, '2001' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('MailServer', 'SRVe', 'Mail Server', 0, NULL, 'mail.unknown.tld' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('MailboxUpcomingEvents', 'SRVe', 'Mailbox Upcoming Events', 0, NULL, 'UpcomingEvents@Unknown.tld' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('MailboxNewsletter', 'SRVe', 'Mailbox Newsletter Notification', 0, NULL, 'NewsletterNotice@Unknown.tld' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('AnnonUser', 'SRVe', 'Notification Email', 0, NULL, 'notifications@unknown.tld' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('AnnonPW', 'SRVe', 'Notification Password', 0, NULL, '' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteZip', 'CFGe', 'Site ZIP (Used for Weather)', 0, 5, '35801' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteChapter', 'CFGe', 'Site Chapter', 0, 5, 'AL-X' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteChapterID', 'CFGe', 'Site Chapter ID', 0, 2, 'X' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteTabNewsletters', 'CFGe', 'Tab Newsletters', 0, NULL, 'Newsletters' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteTabAnnouncements', 'CFGe', 'Tab Announcements', 0, NULL, 'Extra!' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteRSSAnnounceInclude', 'CFGe', 'RSS Announcements (Filter Include)', 0, NULL, '' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('SiteRSSAnnounceExclude', 'CFGe', 'RSS Announcements (Filter Exclude)', 0, NULL, 'calendar,personal,plaque' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('UpcomingEventsSchedule', 'CFGe', 'Email Schedule for Upcoming Events', 0, 8, 'mo,fr' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('MetaKeywords', 'CFGe', 'Homepage Search Keywords', 0, NULL, 'gwrra, goldwing, gold wing road riders association, honda, motorcycle' );

INSERT INTO [config] ([Name], [Access], [Description], [Type], [Maximum], [DataValue])
	VALUES ('CalendarCategories', 'CFGe', 'Homepage Calendar Categories', 0, NULL, '' );







-- -----------------------------------------------------
-- Table [pages]
-- -----------------------------------------------------

CREATE TABLE [pages] (
  [RID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
  [PageName] VARCHAR(64) NOT NULL,
  [SortName] VARCHAR(32) NULL DEFAULT NULL,
  [Format] VARCHAR(4) NULL DEFAULT NULL,
  [Title] VARCHAR(128) NOT NULL,
  [Description] VARCHAR(255) NULL DEFAULT NULL,
  [DateCreated] DATETIME NULL DEFAULT NULL,
  [DateModified] DATETIME NULL DEFAULT NULL,
  [Category] INT NULL DEFAULT NULL,
  [Picture] INT NULL DEFAULT NULL,	-- id of image
  [Disabled] BIT NOT NULL DEFAULT '0',
  [Body] TEXT NULL DEFAULT NULL,
  CONSTRAINT [CategoriesPages]
    FOREIGN KEY ([Category])
    REFERENCES [categories] ([CategoryID])
	ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE INDEX [pagesCategory] ON [pages] ([Category] ASC);
CREATE UNIQUE INDEX [pagesPageName] ON [pages] ([PageName] ASC);



INSERT INTO [pages] ( [PageName], [Format], [Title], [Description], [Body], [DateCreated], [DateModified], [Category]) 
SELECT '~Library', 'STFT', 'Library Image Page (Do Not Delete)', '', '', GETDATE(), GETDATE(), [CategoryID]
FROM [categories] WHERE [Name] = '@Library';

INSERT INTO [pages] ( [PageName], [Format], [Title], [Description], [Body], [DateCreated], [DateModified], [Category]) 
SELECT '@External Pages', 'STFT', 'External Pages (Do Not Delete)', '', '', GETDATE(), GETDATE(), [CategoryID] 
FROM [categories] WHERE [Name] = '@External';

INSERT INTO [pages] ( [PageName], [Format], [Title], [Description], [Body], [DateCreated], [DateModified], [Category]) 
SELECT 'Sample Home Content', 'STFT', 'Sample Home Content', '', 'Sample Home page content', GETDATE(), GETDATE(), [CategoryID] 
FROM [categories] WHERE [Name] = 'Home';

INSERT INTO [pages] ( [PageName], [Format], [Title], [Description], [Body], [DateCreated], [DateModified], [Category]) 
SELECT 'Sample Calendar Content', 'STFT', 'Sample Calendar Content', '', 'Sample Calendar page content', GETDATE(), GETDATE(), [CategoryID]
 FROM [categories] WHERE [Name] = 'Calendar';

INSERT INTO [pages] ( [PageName], [Format], [Title], [Description], [Body], [DateCreated], [DateModified], [Category]) 
SELECT 'Sample Announcement', 'STFT', 'Sample Announcement', 'Stuff for the Description', 'Sample Announcement page content', GETDATE(), GETDATE(), [CategoryID] 
FROM [categories] WHERE [Name] = 'Announcements';

INSERT INTO [pages] ( [PageName], [Format], [Title], [Description], [Body], [DateCreated], [DateModified], [Category]) 
SELECT 'Sample Newsletter', 'STFT', 'Sample Newsletter', NULL, 'Sample Newsletter page content', GETDATE(), GETDATE(), [CategoryID] 
FROM [categories] WHERE [Name] = 'Newsletters';


-- -----------------------------------------------------
-- Table [keywords]
-- -----------------------------------------------------

CREATE TABLE [keywords] (
  [RID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
  [Keyword] VARCHAR(64) NOT NULL,
  [Disabled] BIT NOT NULL DEFAULT '0'
);

CREATE UNIQUE INDEX [keywordsKeyword] ON [keywords] ([Keyword]);


-- -----------------------------------------------------
-- Table [keywordpagemap]
-- -----------------------------------------------------

CREATE TABLE [keywordpagemap] (
  [RID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
  [KeywordID] INT NOT NULL,
  [PageID] INT NOT NULL,
  CONSTRAINT [PagesKeywordPageMap]
    FOREIGN KEY ([PageID])
    REFERENCES [pages] ([RID])
	ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT [KeywordsKeywordPageMap]
    FOREIGN KEY ([KeywordID])
    REFERENCES [keywords] ([RID])
	ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE INDEX [keywordpagemapKeywordID] ON [keywordpagemap] ([KeywordID] ASC);
CREATE INDEX [keywordpagemapPageID] ON [keywordpagemap] ([PageID] ASC);

-- -----------------------------------------------------
-- Table [tabs]
-- -----------------------------------------------------

CREATE TABLE [tabs] (
  [TabID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
  [Name] VARCHAR(50) NULL DEFAULT NULL UNIQUE,
  [SortName] VARCHAR(50) NULL DEFAULT NULL,
  [Description] VARCHAR(255) NULL DEFAULT NULL,
  [Icon] INT NULL DEFAULT NULL,			-- ImageID
  [Picture] INT NULL DEFAULT NULL,		-- ImageID
  [Background] INT NULL DEFAULT NULL,	-- ImageID
  [Disabled] BIT NOT NULL,
);

CREATE UNIQUE INDEX [tabsName] ON [tabs] ([Name] ASC);
CREATE INDEX [tabsSortName] ON [tabs] ([SortName] ASC);


INSERT INTO [tabs] ([Name], [SortName], [Description], [Icon], [Picture], [Background], [Disabled]) 
VALUES ('Home', '000', '@default.asp', NULL, NULL, NULL, 0);

INSERT INTO [tabs] ([Name], [SortName], [Description], [Icon], [Picture], [Background], [Disabled]) 
VALUES ('Forum', NULL, '@forum.asp', NULL, NULL, NULL, 0);

INSERT INTO [tabs] ([Name], [SortName], [Description], [Icon], [Picture], [Background], [Disabled]) 
VALUES ('Calendar', NULL, '@calendar.asp', NULL, NULL, NULL, 0);

INSERT INTO [tabs] ([Name], [SortName], [Description], [Icon], [Picture], [Background], [Disabled]) 
VALUES ('Extra!', NULL, 'Announcements', NULL, NULL, NULL, 0);

INSERT INTO [tabs] ([Name], [SortName], [Description], [Icon], [Picture], [Background], [Disabled]) 
VALUES ('Newsletters', NULL, 'Newsletters', NULL, NULL, NULL, 0);

-- -----------------------------------------------------
-- Table [pagelistoptions]
-- -----------------------------------------------------

CREATE TABLE [pagelistoptions] 
(
  [RID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
  [Name] VARCHAR(64) NOT NULL UNIQUE,
  [ListName] VARCHAR(255) NOT NULL UNIQUE,
  [Details] TEXT NULL DEFAULT NULL
);

INSERT INTO pagelistoptions
	( [Name], ListName, Details )
	VALUES
	( 'CategoryPageASC', 'Category + Page (asc)', 'ISNULL(categories.SortName,'''')+categories.Name, ISNULL(pages.SortName,'''')+pages.PageName' );

INSERT INTO pagelistoptions
	( [Name], ListName, Details )
	VALUES
	( 'CategoryASCPageDESC', 'Category (asc), Page (desc)', 'ISNULL(categories.SortName,'''')+categories.Name, ISNULL(pages.SortName,'''')+pages.PageName desc' );

INSERT INTO pagelistoptions
	( [Name], ListName, Details )
	VALUES
	( 'CategoryPageDESC', 'Category + Page (desc)', 'ISNULL(categories.SortName,'''')+categories.Name desc, ISNULL(pages.SortName,'''')+pages.PageName) desc' );

INSERT INTO pagelistoptions
	( [Name], ListName, Details )
	VALUES
	( 'EventDate', 'Event Date (simple)', 'ISNULL(schedules.DateEvent,GETDATE()+10000), ISNULL(pages.SortName,'''')+pages.PageName' );

INSERT INTO pagelistoptions
	( [Name], ListName, Details )
	VALUES
	( 'EventDateDYN', 'Event Date (Dynamic)', 
 'ISNULL(pages.Sortname,''~''),CASE WHEN (schedules.DateEvent IS NULL) THEN DATEDIFF(day,GETDATE(),pages.DateModified) ELSE dbo.LEAST(DATEDIFF(day,schedules.DateEvent,GETDATE()),DATEDIFF(day,GETDATE(),pages.DateModified)) END,pages.PageName' 
-- 'ISNULL(schedules.DateEvent,GETDATE()+10000), CONCAT(ISNULL(pages.SortName,''),pages.PageName)'
);

INSERT INTO pagelistoptions
	( [Name], ListName, Details )
	VALUES
	( 'DateModifiedDESC', 'Date Modified (desc)', 'pages.DateModified desc, ISNULL(Pages.SortName,'''')+pages.PageName' );


-- -----------------------------------------------------
-- Table [pagelistmap]
-- -----------------------------------------------------

CREATE TABLE [pagelistmap] (
  [RID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
  [TabID] INT NOT NULL,
  [ListID] INT NOT NULL,
  CONSTRAINT [pagelistmapTabID]
    FOREIGN KEY ([TabID])
    REFERENCES [tabs] ([TabID])
	ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT [pagelistmapListID]
    FOREIGN KEY ([ListID])
    REFERENCES [pagelistoptions] ([RID])
	ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE INDEX [pagelistmapRID] ON [pagelistmap] ([RID] ASC);
CREATE INDEX [pagelistmapTabID] ON [pagelistmap] ([TabID] ASC);
CREATE INDEX [PageListOptionspagelistmap] ON [pagelistmap] ([ListID] ASC);

-- INSERT INTO [pagelistmap] ( [TabID], [ListID] )
-- VALUES (
-- 	(SELECT [TabID] FROM [tabs] WHERE [Name]='Extra!'),
-- 	(SELECT [RID] FROM [pagelistoptions] WHERE [Name]='EventDateDYN' )
-- );



-- -----------------------------------------------------
-- Table [images]
-- -----------------------------------------------------

CREATE TABLE [images] (
  [RID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
  [File] VARCHAR(80) NULL DEFAULT NULL,
  [Mime] VARCHAR(40) NULL DEFAULT NULL,
  [DateCreated] DATETIME NOT NULL,
  [Data] VARBINARY(MAX) NULL DEFAULT NULL,
);

CREATE INDEX [imagesRID] ON [images] ([RID]);




-- -----------------------------------------------------
-- Table [pictures]
-- -----------------------------------------------------

CREATE TABLE [pictures] (
  [PictureID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
  [Label] VARCHAR(40) NOT NULL,
  [PageID] INT NOT NULL DEFAULT '0',
  [ImageID] INT NOT NULL,
  CONSTRAINT [PagesPictures]
    FOREIGN KEY ([PageID])
    REFERENCES [pages] ([RID])
	ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT [ImagesPictures]
    FOREIGN KEY ([ImageID])
	REFERENCES [images] ([RID])
	ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE INDEX [picturesLabel] ON [pictures] ([Label] ASC);
CREATE INDEX [picturesPageID] ON [pictures] ([PageID] ASC);

-- -----------------------------------------------------
-- Table [schedules]
-- -----------------------------------------------------

CREATE TABLE [schedules] (
  [RID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
  [PageID] INT NOT NULL,
  [DateBegin] DATETIME NULL DEFAULT NULL,
  [DateEnd] DATETIME NULL DEFAULT NULL,
  [DateEvent] DATETIME NULL DEFAULT NULL,
  [Disabled] BIT NOT NULL DEFAULT '0',
  CONSTRAINT [PagesSchedules]
    FOREIGN KEY ([PageID])
    REFERENCES [pages] ([RID])
	ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE INDEX [schedulesPageID] ON [schedules] ([PageID] ASC);

-- -----------------------------------------------------
-- Table [tabcategorymap]
-- -----------------------------------------------------

CREATE TABLE [tabcategorymap] 
(
  [RID] INT NOT NULL IDENTITY(1,1) PRIMARY KEY,
  [TabID] INT NULL DEFAULT '0',
  [CategoryID] INT NULL DEFAULT '0',
  CONSTRAINT [TabsTabCategoryMap]
    FOREIGN KEY ([TabID])
    REFERENCES [tabs] ([TabID])
	ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT [CategoriesTabCategoryMap]
    FOREIGN KEY ([CategoryID])
    REFERENCES [categories] ([CategoryID])
	ON DELETE CASCADE ON UPDATE CASCADE
);

CREATE INDEX [tabcategorymapCategoryID] ON [tabcategorymap] ([CategoryID] ASC);
CREATE INDEX [tabcategorymapTabID] ON [tabcategorymap] ([TabID] ASC);


INSERT INTO [tabcategorymap] ([TabID], [CategoryID])
SELECT [TabID], [CategoryID]
FROM [tabs], [categories]
WHERE [tabs].[Name] = 'Home' AND [categories].[Name] = '@External';

INSERT INTO [tabcategorymap] ([TabID], [CategoryID])
SELECT [TabID], [CategoryID]
FROM [tabs], [categories]
WHERE [tabs].[Name] = 'Home' AND [categories].[Name] = 'Home';

INSERT INTO [tabcategorymap] ([TabID], [CategoryID])
SELECT [TabID], [CategoryID]
FROM [tabs], [categories]
WHERE [tabs].[Name] = 'Calendar' AND [categories].[Name] = '@External';

INSERT INTO [tabcategorymap] ([TabID], [CategoryID])
SELECT [TabID], [CategoryID]
FROM [tabs], [categories]
WHERE [tabs].[Name] = 'Calendar' AND [categories].[Name] = 'Calendar';

INSERT INTO [tabcategorymap] ([TabID], [CategoryID])
SELECT [TabID], [CategoryID]
FROM [tabs], [categories]
WHERE [tabs].[Name] = 'Forum' AND [categories].[Name] = '@External';

INSERT INTO [tabcategorymap] ([TabID], [CategoryID])
SELECT [TabID], [CategoryID]
FROM [tabs], [categories]
WHERE [tabs].[Name] = 'Extra!' AND [categories].[Name] = 'Announcements';

INSERT INTO [tabcategorymap] ([TabID], [CategoryID])
SELECT [TabID], [CategoryID]
FROM [tabs], [categories]
WHERE [tabs].[Name] = 'Newsletters' AND [categories].[Name] = 'Newsletters';



-- -----------------------------------------------------
-- Table [version]
-- -----------------------------------------------------

CREATE TABLE [version] (
  [Major] INT NOT NULL,
  [Minor] INT NOT NULL,
  [Fixes] INT NULL DEFAULT NULL,
);

INSERT INTO [version]
	( [Major], [Minor], [Fixes] )
	VALUES 
	( 0, 0, 0 );







-- create some resources
--
-- abbrev meanings
--	first three characters are upper case and identify 'real' resource
--	last character is lower case and identifies allowed operation
--	'a' = add/edit
--	'e' = edit
--	'm'	= add/edit/delete

INSERT INTO [authorizeresource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( 'Resource', 'RSCm', 'Maintain Resources', GETDATE() );

INSERT INTO [authorizeresource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( 'Users', 'USRm', 'Maintain Users', GETDATE() );

INSERT INTO [authorizeresource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( 'Passwords', 'PWe', 'May Change Own Password', GETDATE() );

INSERT INTO [authorizeresource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( 'Server', 'SRVe', 'Configure Server Settings', GETDATE() );

INSERT INTO [authorizeresource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( 'SiteConfig', 'CFGe', 'Configure Site Settings', GETDATE() );

INSERT INTO [authorizeresource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( 'Calendar', 'CALm', 'Calendar Edits', GETDATE() );

INSERT INTO [authorizeresource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( 'Navigation', 'NAVm', 'Maintain Navigation Structure', GETDATE() );

INSERT INTO [authorizeresource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( 'PageAuthor', 'PAGm', 'Add and Maintain Pages', GETDATE() );

INSERT INTO [authorizeresource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( 'PageEditor', 'PAGe', 'Edit Page Content', GETDATE() );

INSERT INTO [authorizeresource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( 'NewsAuthor', 'NWSm', 'Add and Maintain Newsletter Pages', GETDATE() );

INSERT INTO [authorizeresource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( 'ChapterMaint', 'CHPm', 'Maintain Chapter Info Pages', GETDATE() );

INSERT INTO [authorizeresource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( 'StyleMaint', 'STYm', 'Maintain and Edit Site Styles', GETDATE() );


-- set up some groups

INSERT INTO [authorizegroup]
	( [Name], [Description], [DateModified] )
	VALUES
	( 'Server Admin', 'Server Administrator with lots of special privs', GETDATE() );

INSERT INTO [authorizegroup]
	( [Name], [Description], [DateModified] )
	VALUES
	( 'Webmaster', 'Webmaster responsible for most of the maintenance of the web content', GETDATE() );

INSERT INTO [authorizegroup]
	( [Name], [Description], [DateModified] )
	VALUES
	( 'Newsletter Editor', 'Responsible for adding content to the newsletter', GETDATE() );

INSERT INTO [authorizegroup]
	( [Name], [Description], [DateModified] )
	VALUES
	( 'Site Editor', 'Responsible for checking content on the website', GETDATE() );

INSERT INTO [authorizegroup]
	( [Name], [Description], [DateModified] )
	VALUES
	( 'Calendar Editor', 'Responsible for Calendar Edits', GETDATE() );

INSERT INTO [authorizegroup]
	( [Name], [Description], [DateModified] )
	VALUES
	( 'Chapter Editor', 'Responsible for maintaining Chapter Info pages', GETDATE() );


-- map the resources into the groups

-- Server Admin

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'Resource' )
	FROM [authorizegroup]
	WHERE [Name] = 'Server Admin';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'Users' )
	FROM [authorizegroup]
	WHERE [Name] = 'Server Admin';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'Passwords' )
	FROM [authorizegroup]
	WHERE [Name] = 'Server Admin';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'Server' )
	FROM [authorizegroup]
	WHERE [Name] = 'Server Admin';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'SiteConfig' )
	FROM [authorizegroup]
	WHERE [Name] = 'Server Admin';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'Calendar' )
	FROM [authorizegroup]
	WHERE [Name] = 'Server Admin';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'Navigation' )
	FROM [authorizegroup]
	WHERE [Name] = 'Server Admin';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'PageAuthor' )
	FROM [authorizegroup]
	WHERE [Name] = 'Server Admin';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'ChapterMaint' )
	FROM [authorizegroup]
	WHERE [Name] = 'Server Admin';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'StyleMaint' )
	FROM [authorizegroup]
	WHERE [Name] = 'Server Admin';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'NewsAuthor' )
	FROM [authorizegroup]
	WHERE [Name] = 'Server Admin';


-- Webmaster

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'Passwords' )
	FROM [authorizegroup]
	WHERE [Name] = 'Webmaster';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'SiteConfig' )
	FROM [authorizegroup]
	WHERE [Name] = 'Webmaster';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'Calendar' )
	FROM [authorizegroup]
	WHERE [Name] = 'Webmaster';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'Navigation' )
	FROM [authorizegroup]
	WHERE [Name] = 'Webmaster';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'PageAuthor' )
	FROM [authorizegroup]
	WHERE [Name] = 'Webmaster';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'ChapterMaint' )
	FROM [authorizegroup]
	WHERE [Name] = 'Webmaster';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'NewsAuthor' )
	FROM [authorizegroup]
	WHERE [Name] = 'Webmaster';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'StyleMaint' )
	FROM [authorizegroup]
	WHERE [Name] = 'Webmaster';



-- Newsletter Editor

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'Passwords' )
	FROM [authorizegroup]
	WHERE [Name] = 'Newsletter Editor';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'NewsAuthor' )
	FROM [authorizegroup]
	WHERE [Name] = 'Newsletter Editor';



-- Site Editor

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'Passwords' )
	FROM [authorizegroup]
	WHERE [Name] = 'Site Editor';

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'PageEditor' )
	FROM [authorizegroup]
	WHERE [Name] = 'Site Editor';


-- Calendar Editor

INSERT INTO [authorizeresourcemap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [authorizeresource] WHERE [Name] = 'Calendar' )
	FROM [authorizegroup]
	WHERE [Name] = 'Calendar Editor';



-- now let's map the groups to a user

INSERT INTO [authorizegroupmap]
	( [UserID], [GroupID] )
	SELECT [RID], (SELECT [RID] FROM [authorizegroup] WHERE [Name] = 'Server Admin')
	FROM [authorizeusers]
	WHERE [Username] = 'grizzie17';

INSERT INTO [authorizegroupmap]
	( [UserID], [GroupID] )
	SELECT [RID], (SELECT [RID] FROM [authorizegroup] WHERE [Name] = 'Webmaster')
	FROM [authorizeusers]
	WHERE [Username] = 'grizzie17';

INSERT INTO [authorizegroupmap]
	( [UserID], [GroupID] )
	SELECT [RID], (SELECT [RID] FROM [authorizegroup] WHERE [Name] = 'Webmaster')
	FROM [authorizeusers]
	WHERE [Username] = 'webby17';





