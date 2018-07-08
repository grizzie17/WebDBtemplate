# alter the Config table


ALTER TABLE [Config]
	ADD COLUMN [SiteZip] TEXT(12)  WITH COMPRESSION NULL;

ALTER TABLE [Config]
	ADD COLUMN [SiteChapter] TEXT(5)  WITH COMPRESSION NULL;

ALTER TABLE [Config]
	ADD COLUMN [SiteChapterID] TEXT(2)  WITH COMPRESSION NULL;

ALTER TABLE [Config]
	ADD COLUMN [SiteTabNewsletters] TEXT(64)  WITH COMPRESSION NULL;

UPDATE [Config]
	Set [SiteTabNewsletters] = "Newsletters"
	WHERE
		SiteTabNewsletters IS NULL;


# next 3 statements accomplish a rename of HomepageAnnouncements to SiteTabAnnouncements

ALTER TABLE [Config]
	ADD COLUMN [SiteTabAnnouncements] TEXT(64)  WITH COMPRESSION NULL;

UPDATE [Config]
	SET [SiteTabAnnouncements] = [HomepageAnnouncements]
	WHERE
		NOT HomepageAnnouncements IS NULL;

ALTER TABLE [Config]
	DROP COLUMN [HomepageAnnouncements];



UPDATE Version
	SET Major = 4, Minor = 1, Fixes = 0;

