# Add some stuff for the Users table

ALTER TABLE [AuthorizeUsers]
	ADD COLUMN [Lastname] TEXT(32)  WITH COMPRESSION NULL;

ALTER TABLE [AuthorizeUsers]
	ADD COLUMN [Firstname] TEXT(32)  WITH COMPRESSION NULL;


# create resources that we can control access to

CREATE TABLE AuthorizeResource
(
	[RID] AUTOINCREMENT NOT NULL,
	[Name] TEXT(40) WITH COMPRESSION NOT NULL,
	[Abbrev]  TEXT(4) WITH COMPRESSION NOT NULL,
	[Description] TEXT(255) WITH COMPRESSION NULL,
	[DateModified] DATETIME  NOT NULL,
	[Disabled] YESNO  DEFAULT 0,
	CONSTRAINT [PrimaryKey] PRIMARY KEY( [RID] ),
	CONSTRAINT [RID] UNIQUE ([RID])
);

CREATE INDEX [AuthorizeResourceRID] ON [AuthorizeResource]([RID]);
CREATE INDEX [AuthorizeResourceName] ON [AuthorizeResource]([Name]);


# group resources into nice digestible chunks

CREATE TABLE AuthorizeGroup
(
	[RID] AUTOINCREMENT NOT NULL,
	[Name] TEXT(40) WITH COMPRESSION NOT NULL,
	[Description] TEXT(255) WITH COMPRESSION NULL,
	[DateModified] DATETIME  NOT NULL,
	[Disabled] YESNO  DEFAULT 0,
	CONSTRAINT [PrimaryKey] PRIMARY KEY( [RID] ),
	CONSTRAINT [RID] UNIQUE ([RID])
);

CREATE INDEX [AuthorizeGroupRID] ON [AuthorizeGroup]([RID]);
CREATE INDEX [AuthorizeGroupName] ON [AuthorizeGroup]([Name]);


CREATE TABLE AuthorizeResourceMap
(
	[RID] AUTOINCREMENT  NOT NULL,
	[GroupID] LONG  NOT NULL,
	[ResourceID] LONG  NOT NULL,
	CONSTRAINT [PrimaryKey] PRIMARY KEY( [RID] )
);

CREATE INDEX [AuthorizeResourceMapGroupID] ON [AuthorizeResourceMap]([GroupID]);
CREATE INDEX [AuthorizeResourceMapRID] ON [AuthorizeResourceMap]([RID]);

#note: FOREIGN KEYs are always One-to-Many relationships

ALTER TABLE [AuthorizeResourceMap]
	ADD CONSTRAINT [GroupsAuthorizeResourceMap]
	FOREIGN KEY NO INDEX ([GroupID]) 
		REFERENCES [AuthorizeGroup] ([RID]);
ALTER TABLE [AuthorizeResourceMap]
	ADD CONSTRAINT [ResourcesAuthorizeResourceMap]
	FOREIGN KEY NO INDEX ([ResourceID]) 
		REFERENCES [AuthorizeResource] ([RID]);



# Now associate the groups to users

CREATE TABLE AuthorizeGroupMap
(
	[RID] AUTOINCREMENT  NOT NULL,
	[UserID] LONG  NOT NULL,
	[GroupID] LONG  NOT NULL,
	CONSTRAINT [PrimaryKey] PRIMARY KEY( [RID] )
);

CREATE INDEX [AuthorizeGroupMapUserID] ON [AuthorizeGroupMap]([UserID]);
CREATE INDEX [AuthorizeGroupMapRID] ON [AuthorizeGroupMap]([RID]);

ALTER TABLE [AuthorizeGroupMap]
	ADD CONSTRAINT [UsersAuthorizeGroupMap]
	FOREIGN KEY NO INDEX ([UserID]) 
		REFERENCES [AuthorizeUsers] ([RID]);
ALTER TABLE [AuthorizeGroupMap]
	ADD CONSTRAINT [GroupsAuthorizeGroupMap]
	FOREIGN KEY NO INDEX ([GroupID]) 
		REFERENCES [AuthorizeGroup] ([RID]);


# table to support remembering user

CREATE TABLE AuthorizeRemember
(
	[RID] AUTOINCREMENT  NOT NULL,
	[UserID]  LONG  NOT NULL,
	[Cookie]  TEXT(40)  WITH COMPRESSION NOT NULL,
	[DateModified] DATETIME  NOT NULL,
	CONSTRAINT [PrimaryKey] PRIMARY KEY( [RID] ),
	CONSTRAINT [RID] UNIQUE ([RID])
);

CREATE INDEX [AuthorizeRememberUserID] ON [AuthorizeRemember]([UserID]);
CREATE INDEX [AuthorizeRememberRID] ON [AuthorizeRemember]([RID]);

ALTER TABLE [AuthorizeRemember]
	ADD CONSTRAINT [UsersAuthorizeRemember]
	FOREIGN KEY NO INDEX ([UserID]) 
		REFERENCES [AuthorizeUsers] ([RID]);


#
#
# now populate with some data
#

# first let's use the name stuff

UPDATE [AuthorizeUsers]
	SET [Lastname] = "Griswold", [Firstname] = "John", [DateModified] = #1/9/2009#
	WHERE [Username] = "grizzie17";


# create some resources
#
# abbrev meanings
#	first three characters are upper case and identify "real" resource
#	last character is lower case and identifies allowed operation
#	"a" = add/edit
#	"e" = edit
#	"m"	= add/edit/delete

INSERT INTO [AuthorizeResource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( "Resource", "RSCm", "Maintain Resources", #1/9/2009# );

INSERT INTO [AuthorizeResource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( "Users", "USRm", "Maintain Users", #1/9/2009# );

INSERT INTO [AuthorizeResource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( "Passwords", "PWe", "May Change Own Password", #1/9/2009# );

INSERT INTO [AuthorizeResource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( "Server", "SRVe", "Configure Server Settings", #1/9/2009# );

INSERT INTO [AuthorizeResource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( "SiteConfig", "CFGe", "Configure Site Settings", #1/9/2009# );

INSERT INTO [AuthorizeResource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( "Calendar", "CALm", "Calendar Edits", #1/9/2009# );

INSERT INTO [AuthorizeResource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( "Navigation", "NAVm", "Maintain Navigation Structure", #1/9/2009# );

INSERT INTO [AuthorizeResource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( "PageAuthor", "PAGm", "Add and Maintain Pages", #1/9/2009# );

INSERT INTO [AuthorizeResource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( "PageEditor", "PAGe", "Edit Page Content", #1/9/2009# );

INSERT INTO [AuthorizeResource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( "NewsAuthor", "NWSm", "Add and Maintain Newsletter Pages", #1/9/2009# );


# set up some groups

INSERT INTO [AuthorizeGroup]
	( [Name], [Description], [DateModified] )
	VALUES
	( "Server Admin", "Server Administrator with lots of special privs", #1/9/2009# );

INSERT INTO [AuthorizeGroup]
	( [Name], [Description], [DateModified] )
	VALUES
	( "Webmaster", "Webmaster responsible for most of the maintenance of the web content", #1/9/2009# );

INSERT INTO [AuthorizeGroup]
	( [Name], [Description], [DateModified] )
	VALUES
	( "Newsletter Editor", "Responsible for adding content to the newsletter", #1/9/2009# );

INSERT INTO [AuthorizeGroup]
	( [Name], [Description], [DateModified] )
	VALUES
	( "Site Editor", "Responsible for checking content on the website", #1/9/2009# );


# map the resources into the groups

## Server Admin

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "Resource" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Server Admin";

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "Users" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Server Admin";

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "Passwords" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Server Admin";

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "Server" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Server Admin";

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "SiteConfig" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Server Admin";

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "Calendar" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Server Admin";

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "Navigation" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Server Admin";

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "PageAuthor" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Server Admin";

## Webmaster

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "Passwords" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Webmaster";

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "SiteConfig" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Webmaster";

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "Calendar" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Webmaster";

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "Navigation" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Webmaster";

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "PageAuthor" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Webmaster";


## Newsletter Editor

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "Passwords" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Newsletter Editor";

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "NewsAuthor" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Newsletter Editor";



## Site Editor

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "Passwords" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Site Editor";

INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "PageEditor" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Site Editor";



#
# now let's map the groups to a user

INSERT INTO [AuthorizeGroupMap]
	( [UserID], [GroupID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeGroup] WHERE [Name] = "Server Admin")
	FROM [AuthorizeUsers]
	WHERE [Username] = "grizzie17";

INSERT INTO [AuthorizeGroupMap]
	( [UserID], [GroupID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeGroup] WHERE [Name] = "Webmaster")
	FROM [AuthorizeUsers]
	WHERE [Username] = "grizzie17";









UPDATE [Version]
	SET Major = 7, Minor = 0, Fixes = 0;

