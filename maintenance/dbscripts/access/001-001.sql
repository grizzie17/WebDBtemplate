# add the Config table

CREATE TABLE [Config]
(
	[RID] AUTOINCREMENT  NOT NULL,
	[Domain] TEXT(128) WITH COMPRESSION NOT NULL,
	[SiteName] TEXT(255) WITH COMPRESSION NOT NULL,
	[SiteTitle] TEXT(255) WITH COMPRESSION NULL,
	[SiteOrgDesignation] TEXT(128) WITH COMPRESSION NULL,
	[SiteMotto] TEXT(255) WITH COMPRESSION NULL,
	[SiteLocation] TEXT(128) WITH COMPRESSION NULL,
	[CopyrightStartYear] INTEGER,
	[MailServer] TEXT(128) WITH COMPRESSION,
	[MailboxUpcomingEvents] Text(128) WITH COMPRESSION,
	[MailboxNewsletter] Text(128) WITH COMPRESSION,
	[MailUser] Text(128) WITH COMPRESSION,
	[MailPW] Text(64) WITH COMPRESSION,
	[HomepageAnnouncements] TEXT(128) WITH COMPRESSION NOT NULL,
	CONSTRAINT [PrimaryKey] PRIMARY KEY ([RID])
);


INSERT INTO [Config]
	( [Domain], 
	[SiteName], [SiteTitle], [SiteOrgDesignation], 
	[SiteMotto], 
	[SiteLocation], [CopyrightStartYear], 
	[MailServer], [MailboxUpcomingEvents], [MailboxNewsletter],
	[MailUser], [MailPW],
	[HomepageAnnouncements] )
	VALUES
	( "Unknown.tld",
	"Unknown Site", NULL, "Chapter XX-?",
	"This is Your Motto!",
	"Anywhereville, XX", 0,
	"mail.unknown.tld", "UpcomingEvents@Unknown.tld", "NewsletterNotice@Unknown.tld",
	"notifications@unknown.tld", NULL,
	"Extra!" );




UPDATE Version
	SET Major = 2, Minor = 0, Fixes = 0;
