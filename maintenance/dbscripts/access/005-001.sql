# New tables for user permissions


CREATE TABLE AuthorizeUsers
(
	[RID] AUTOINCREMENT NOT NULL,
	[Username] TEXT(32) WITH COMPRESSION NOT NULL,
	[Password] TEXT(32) WITH COMPRESSION NOT NULL,
	[Email] TEXT(128) WITH COMPRESSION NOT NULL,
	[Description] TEXT(255) WITH COMPRESSION NULL,
	[DateCreated] DATETIME  NOT NULL,
	[DateModified] DATETIME  NOT NULL,
	[Disabled] YESNO  DEFAULT 0,
	CONSTRAINT [PrimaryKey] PRIMARY KEY( [RID] ),
	CONSTRAINT [RID] UNIQUE ([RID])
);

CREATE INDEX [AuthorizeUsersRID] ON [AuthorizeUsers]([RID]);
CREATE INDEX [AuthorizeUsersUsername] ON [AuthorizeUsers]([Username]);


CREATE TABLE AuthorizeLog
(
	[UserID] LONG  NOT NULL,
	[DateLogon]  DATETIME  NOT NULL,
	[DateFail]  DATETIME  NOT NULL,
	[FailCount]  LONG  NOT NULL,
	CONSTRAINT [PrimaryKey] PRIMARY KEY( [UserID] ),
	CONSTRAINT [UserID] UNIQUE ([UserID])
);

#note: FOREIGN KEYs are always One-to-Many relationships

ALTER TABLE [AuthorizeLog]
	ADD CONSTRAINT [AuthorizeUsersAuthorizeLog]
	FOREIGN KEY NO INDEX ([UserID]) 
		REFERENCES [AuthorizeUsers] ([RID]);



INSERT INTO [AuthorizeUsers]
	( [UserName], [Password], [Email], [DateCreated], [DateModified] )
	VALUES
	( "grizzie17", "grizzly17", "john@grizzlyweb.com", #1/1/2009#, #1/1/2009# );

INSERT INTO [AuthorizeUsers]
	( [UserName], [Password], [Email], [DateCreated], [DateModified] )
	VALUES
	( "admin123", "123.admin!321", "", #1/1/2009#, #1/1/2009# );



#note: in the future we need to add tables for handling different types of authorization


UPDATE Version
	SET Major = 6, Minor = 0, Fixes = 0;

