/* Access SQL export data follows. Auto-generated. */

/* Tables */
-- -----------------------------------------------------
-- Table [authorizeusers]
-- -----------------------------------------------------

DROP TABLE [AuthorizeUsers] /**WEAK**/;
CREATE TABLE [AuthorizeUsers] (
[RID] COUNTER NOT NULL,
[Username] TEXT(32) NOT NULL,
[Password] TEXT(32) NOT NULL,
[Email] TEXT(128) NOT NULL,
[Description] TEXT(255),
[DateCreated] DATETIME NOT NULL,
[DateModified] DATETIME NOT NULL,
[Disabled] BIT NOT NULL DEFAULT 0,
[Lastname] TEXT(32),
[Firstname] TEXT(32)
);
CREATE UNIQUE INDEX [PrimaryKey] ON [AuthorizeUsers] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE UNIQUE INDEX [RID] ON [AuthorizeUsers] (RID ASC);
CREATE INDEX [AuthorizeUsersRID] ON [AuthorizeUsers] (RID ASC);
CREATE INDEX [AuthorizeUsersUsername] ON [AuthorizeUsers] (Username ASC);

-- -----------------------------------------------------
-- Table [authorizegroup]
-- -----------------------------------------------------

DROP TABLE [AuthorizeGroup] /**WEAK**/;
CREATE TABLE [AuthorizeGroup] (
[RID] COUNTER NOT NULL,
[Name] TEXT(40) NOT NULL,
[Description] TEXT(255),
[DateModified] DATETIME NOT NULL,
[Disabled] BIT NOT NULL DEFAULT 0
);
CREATE UNIQUE INDEX [PrimaryKey] ON [AuthorizeGroup] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE UNIQUE INDEX [RID] ON [AuthorizeGroup] (RID ASC);
CREATE INDEX [AuthorizeGroupName] ON [AuthorizeGroup] (Name ASC);
CREATE INDEX [AuthorizeGroupRID] ON [AuthorizeGroup] (RID ASC);

-- -----------------------------------------------------
-- Table [authorizegroupmap]
-- -----------------------------------------------------

DROP TABLE [AuthorizeGroupMap] /**WEAK**/;
CREATE TABLE [AuthorizeGroupMap] (
[RID] COUNTER NOT NULL,
[UserID] INTEGER NOT NULL,
[GroupID] INTEGER NOT NULL
);
CREATE UNIQUE INDEX [PrimaryKey] ON [AuthorizeGroupMap] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE INDEX [AuthorizeGroupMapRID] ON [AuthorizeGroupMap] (RID ASC);
CREATE INDEX [AuthorizeGroupMapUserID] ON [AuthorizeGroupMap] (UserID ASC);

-- -----------------------------------------------------
-- Table [authorizelog]
-- -----------------------------------------------------

DROP TABLE [AuthorizeLog] /**WEAK**/;
CREATE TABLE [AuthorizeLog] (
[UserID] INTEGER NOT NULL,
[DateLogon] DATETIME NOT NULL,
[DateFail] DATETIME NOT NULL,
[FailCount] INTEGER NOT NULL
);
CREATE UNIQUE INDEX [PrimaryKey] ON [AuthorizeLog] (UserID ASC) WITH PRIMARY DISALLOW NULL;
CREATE UNIQUE INDEX [UserID] ON [AuthorizeLog] (UserID ASC);

-- -----------------------------------------------------
-- Table [authorizeremember]
-- -----------------------------------------------------

DROP TABLE [AuthorizeRemember] /**WEAK**/;
CREATE TABLE [AuthorizeRemember] (
[RID] COUNTER NOT NULL,
[UserID] INTEGER NOT NULL,
[Cookie] TEXT(40) NOT NULL,
[DateModified] DATETIME NOT NULL
);
CREATE UNIQUE INDEX [PrimaryKey] ON [AuthorizeRemember] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE UNIQUE INDEX [RID] ON [AuthorizeRemember] (RID ASC);
CREATE INDEX [AuthorizeRememberRID] ON [AuthorizeRemember] (RID ASC);
CREATE INDEX [AuthorizeRememberUserID] ON [AuthorizeRemember] (UserID ASC);

-- -----------------------------------------------------
-- Table [authorizeresource]
-- -----------------------------------------------------

DROP TABLE [AuthorizeResource] /**WEAK**/;
CREATE TABLE [AuthorizeResource] (
[RID] COUNTER NOT NULL,
[Name] TEXT(40) NOT NULL,
[Abbrev] TEXT(4) NOT NULL,
[Description] TEXT(255),
[DateModified] DATETIME NOT NULL,
[Disabled] BIT NOT NULL DEFAULT 0
);
CREATE UNIQUE INDEX [PrimaryKey] ON [AuthorizeResource] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE UNIQUE INDEX [RID] ON [AuthorizeResource] (RID ASC);
CREATE INDEX [AuthorizeResourceName] ON [AuthorizeResource] (Name ASC);
CREATE INDEX [AuthorizeResourceRID] ON [AuthorizeResource] (RID ASC);

-- -----------------------------------------------------
-- Table [authorizeresourcemap]
-- -----------------------------------------------------

DROP TABLE [AuthorizeResourceMap] /**WEAK**/;
CREATE TABLE [AuthorizeResourceMap] (
[RID] COUNTER NOT NULL,
[GroupID] INTEGER NOT NULL,
[ResourceID] INTEGER NOT NULL
);
CREATE UNIQUE INDEX [PrimaryKey] ON [AuthorizeResourceMap] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE INDEX [AuthorizeResourceMapGroupID] ON [AuthorizeResourceMap] (GroupID ASC);
CREATE INDEX [AuthorizeResourceMapRID] ON [AuthorizeResourceMap] (RID ASC);

-- -----------------------------------------------------
-- Table [categories]
-- -----------------------------------------------------

DROP TABLE [Categories] /**WEAK**/;
CREATE TABLE [Categories] (
[CategoryID] COUNTER NOT NULL,
[Name] TEXT(50) NOT NULL,
[ShortName] TEXT(30) /**COMMENT* Used when dealing with the "tab" category view */,
[ShortDesc] LONGTEXT,
[Icon] TEXT(40) /**COMMENT* Relative URL/name of a picture that can be used as a bullet for the list of categories */,
[Disabled] BIT NOT NULL DEFAULT 0,
[SortName] TEXT(20) /**Increase to 32 */
/*[Revision] INTEGER NOT NULL DEFAULT 0 */
);
CREATE UNIQUE INDEX [CategoryID] ON [Categories] (CategoryID ASC);
CREATE UNIQUE INDEX [PrimaryKey] ON [Categories] (CategoryID ASC) WITH PRIMARY DISALLOW NULL;

-- -----------------------------------------------------
-- Table [config]
-- -----------------------------------------------------
/** COMPLETE REWORK **/
DROP TABLE [Config] /**WEAK**/;
CREATE TABLE [Config] (
[RID] COUNTER NOT NULL,
[Domain] TEXT(128) NOT NULL,
[SiteName] TEXT(255) NOT NULL,
[SiteTitle] TEXT(255),
[SiteOrgDesignation] TEXT(128),
[SiteMotto] TEXT(255),
[SiteLocation] TEXT(128),
[CopyrightStartYear] INTEGER,
[MailServer] TEXT(128),
[MailboxUpcomingEvents] TEXT(128),
[MailboxNewsletter] TEXT(128),
[MailUser] TEXT(128),
[MailPW] TEXT(64),
[SiteZip] TEXT(12),
[SiteChapter] TEXT(5),
[SiteChapterID] TEXT(2),
[SiteTabNewsletters] TEXT(64),
[SiteTabAnnouncements] TEXT(64),
[SiteRSSAnnounceInclude] TEXT(64),
[SiteRSSAnnounceExclude] TEXT(64)
);
CREATE UNIQUE INDEX [PrimaryKey] ON [Config] (RID ASC) WITH PRIMARY DISALLOW NULL;





-- -----------------------------------------------------
-- Table [pages]
-- -----------------------------------------------------

DROP TABLE [Pages] /**WEAK**/;
CREATE TABLE [Pages] (
[RID] COUNTER NOT NULL,
[PageName] TEXT(64) NOT NULL,
[SortName] TEXT(32),
[Format] TEXT(4) /**COMMENT* Format of Title and Body: TEXT, HTML, STFT (Simple Text Format) */,
[Title] TEXT(128) NOT NULL,
[Description] TEXT(255),
[Body] LONGTEXT,
[DateCreated] DATETIME,
[DateModified] DATETIME,
[Category] INTEGER,
[Picture] TEXT(40), /** change to number **/
[Disabled] BIT NOT NULL DEFAULT 0
);
CREATE UNIQUE INDEX [PrimaryKey] ON [Pages] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE UNIQUE INDEX [RID] ON [Pages] (RID ASC);
CREATE INDEX [Category] ON [Pages] (Category ASC);
CREATE INDEX [PageName] ON [Pages] (PageName ASC);


-- -----------------------------------------------------
-- Table [keywords]
-- -----------------------------------------------------

DROP TABLE [Keywords] /**WEAK**/;
CREATE TABLE [Keywords] (
[RID] COUNTER NOT NULL,
[Keyword] TEXT(64) NOT NULL,
[Disabled] BIT NOT NULL DEFAULT 0
);
CREATE UNIQUE INDEX [Keyword] ON [Keywords] (Keyword ASC);
CREATE UNIQUE INDEX [PrimaryKey] ON [Keywords] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE UNIQUE INDEX [RID] ON [Keywords] (RID ASC);


-- -----------------------------------------------------
-- Table [keywordpagemap]
-- -----------------------------------------------------

DROP TABLE [KeywordPageMap] /**WEAK**/;
CREATE TABLE [KeywordPageMap] (
[RID] COUNTER NOT NULL,
[KeywordID] INTEGER NOT NULL,
[PageID] INTEGER NOT NULL
);
CREATE UNIQUE INDEX [PrimaryKey] ON [KeywordPageMap] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE UNIQUE INDEX [RID] ON [KeywordPageMap] (RID ASC);
CREATE INDEX [KeywordID] ON [KeywordPageMap] (KeywordID ASC);
CREATE INDEX [PageID] ON [KeywordPageMap] (PageID ASC);


-- -----------------------------------------------------
-- Table [tabs]
-- -----------------------------------------------------

DROP TABLE [Tabs] /**WEAK**/;
CREATE TABLE [Tabs] (
[TabID] COUNTER NOT NULL,
[Name] TEXT(50),
[SortName] TEXT(50),
[Description] LONGTEXT,
[Icon] TEXT(40) /**Change to imageID**/ /**COMMENT* Optional picture to be displayed on the Tab itself */,
[Picture] TEXT(40)  /**Change to ImageID**/ /**COMMENT* URL of picture that is to be displayed in the tab body */,
[Background] TEXT(40)   /**ImageID**/ /**COMMENT* URL of picture to be used as page background */,
[Disabled] BIT NOT NULL
);
CREATE UNIQUE INDEX [PrimaryKey] ON [Tabs] (TabID ASC) WITH PRIMARY DISALLOW NULL;
CREATE UNIQUE INDEX [TabID] ON [Tabs] (TabID ASC);
CREATE INDEX [SortName] ON [Tabs] (SortName ASC);


-- -----------------------------------------------------
-- Table [pagelistoptions]
-- -----------------------------------------------------

DROP TABLE [PageListOptions] /**WEAK**/;
CREATE TABLE [PageListOptions] (
[RID] COUNTER NOT NULL,
/**[Name] TEXT(64) NOT NULL,**/
[ListName] TEXT(255) NOT NULL,
[Details] LONGTEXT
);
CREATE UNIQUE INDEX [PrimaryKey] ON [PageListOptions] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE UNIQUE INDEX [RID] ON [PageListOptions] (RID ASC);


-- -----------------------------------------------------
-- Table [pagelistmap]
-- -----------------------------------------------------

DROP TABLE [PageListMap] /**WEAK**/;
CREATE TABLE [PageListMap] (
[RID] COUNTER NOT NULL,
[TabID] INTEGER NOT NULL,
[ListID] INTEGER NOT NULL
);
CREATE UNIQUE INDEX [PrimaryKey] ON [PageListMap] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE INDEX [PageListMapTabID] ON [PageListMap] (TabID ASC);
CREATE INDEX [RID] ON [PageListMap] (RID ASC);


-- -----------------------------------------------------
-- Table [images]
-- -----------------------------------------------------
/**NEW TABLE**/
CREATE TABLE [images] (
    [RID] COUNTER NOT NULL,
    [File] TEXT(80),
    [Mime] TEXT(40),
    [DateCreated] DATE NOT NULL,
);

-- -----------------------------------------------------
-- Table [pictures]
-- -----------------------------------------------------

DROP TABLE [Pictures] /**WEAK**/;
CREATE TABLE [Pictures] (
[PictureID] COUNTER NOT NULL,
[PageID] INTEGER NOT NULL DEFAULT 0,
[Label] TEXT(40) NOT NULL,
[ImageID] INTEGER NOT NULL
[File] TEXT(40) /**DROP**/
);
CREATE UNIQUE INDEX [PictureID] ON [Pictures] (PictureID ASC);
CREATE UNIQUE INDEX [PrimaryKey] ON [Pictures] (PictureID ASC) WITH PRIMARY DISALLOW NULL;
CREATE INDEX [Label] ON [Pictures] (Label ASC);
CREATE INDEX [ProductID] ON [Pictures] (PageID ASC);


-- -----------------------------------------------------
-- Table [schedules]
-- -----------------------------------------------------

DROP TABLE [Schedules] /**WEAK**/;
CREATE TABLE [Schedules] (
[RID] COUNTER NOT NULL,
[PageID] INTEGER NOT NULL,
[DateBegin] DATETIME,
[DateEnd] DATETIME,
[DateEvent] DATETIME,
[Disabled] BIT NOT NULL DEFAULT 0
);
CREATE UNIQUE INDEX [PrimaryKey] ON [Schedules] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE UNIQUE INDEX [RID] ON [Schedules] (RID ASC);
CREATE INDEX [PageID] ON [Schedules] (PageID ASC);


-- -----------------------------------------------------
-- Table [tabcategorymap]
-- -----------------------------------------------------

DROP TABLE [TabCategoryMap] /**WEAK**/;
CREATE TABLE [TabCategoryMap] (
[RID] COUNTER NOT NULL,
[TabID] INTEGER DEFAULT 0,
[CategoryID] INTEGER DEFAULT 0
);
CREATE UNIQUE INDEX [PrimaryKey] ON [TabCategoryMap] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE UNIQUE INDEX [RID] ON [TabCategoryMap] (RID ASC);
CREATE INDEX [CategoryID] ON [TabCategoryMap] (CategoryID ASC);
CREATE INDEX [TabID] ON [TabCategoryMap] (TabID ASC);


-- -----------------------------------------------------
-- Table [version]
-- -----------------------------------------------------

DROP TABLE [Version] /**WEAK**/;
CREATE TABLE [Version] (
[Major] INTEGER NOT NULL,
[Minor] INTEGER NOT NULL,
[Fixes] INTEGER
);
CREATE UNIQUE INDEX [Index_5790E1A2_C047_4C26] ON [Version] (Major ASC) WITH PRIMARY DISALLOW NULL;

/* Foreign keys for AuthorizeGroupMap */
ALTER TABLE [AuthorizeGroupMap] ADD CONSTRAINT [GroupsAuthorizeGroupMap] FOREIGN KEY (GroupID) REFERENCES [AuthorizeGroup] (RID);
ALTER TABLE [AuthorizeGroupMap] ADD CONSTRAINT [UsersAuthorizeGroupMap] FOREIGN KEY (UserID) REFERENCES [AuthorizeUsers] (RID);

/* Foreign keys for AuthorizeLog */
ALTER TABLE [AuthorizeLog] ADD CONSTRAINT [AuthorizeUsersAuthorizeLog] FOREIGN KEY (UserID) REFERENCES [AuthorizeUsers] (RID);

/* Foreign keys for AuthorizeRemember */
ALTER TABLE [AuthorizeRemember] ADD CONSTRAINT [UsersAuthorizeRemember] FOREIGN KEY (UserID) REFERENCES [AuthorizeUsers] (RID);

/* Foreign keys for AuthorizeResourceMap */
ALTER TABLE [AuthorizeResourceMap] ADD CONSTRAINT [GroupsAuthorizeResourceMap] FOREIGN KEY (GroupID) REFERENCES [AuthorizeGroup] (RID);
ALTER TABLE [AuthorizeResourceMap] ADD CONSTRAINT [ResourcesAuthorizeResourceMap] FOREIGN KEY (ResourceID) REFERENCES [AuthorizeResource] (RID);

/* Foreign keys for Gallery */
ALTER TABLE [Gallery] ADD CONSTRAINT [GalleryPicturesGallery] FOREIGN KEY (CoverPictureID) REFERENCES [GalleryPictures] (RID);

/* Foreign keys for GalleryCategoryMap */
ALTER TABLE [GalleryCategoryMap] ADD CONSTRAINT [GalleryCategoriesGalleryCategoryMap] FOREIGN KEY (CategoryID) REFERENCES [GalleryCategories] (RID);
ALTER TABLE [GalleryCategoryMap] ADD CONSTRAINT [GalleryGalleryCategoryMap] FOREIGN KEY (GalleryID) REFERENCES [Gallery] (RID);

/* Foreign keys for GalleryPictures */
ALTER TABLE [GalleryPictures] ADD CONSTRAINT [GalleryGalleryPictures] FOREIGN KEY (GalleryID) REFERENCES [Gallery] (RID);

/* Foreign keys for GalleryShuffle */
ALTER TABLE [GalleryShuffle] ADD CONSTRAINT [GalleryPicturesGalleryShuffle] FOREIGN KEY (PictureID) REFERENCES [GalleryPictures] (RID);

/* Foreign keys for KeywordPageMap */
ALTER TABLE [KeywordPageMap] ADD CONSTRAINT [KeywordsKeywordPageMap] FOREIGN KEY (KeywordID) REFERENCES [Keywords] (RID);
ALTER TABLE [KeywordPageMap] ADD CONSTRAINT [PagesKeywordPageMap] FOREIGN KEY (PageID) REFERENCES [Pages] (RID);

/* Foreign keys for PageListMap */
ALTER TABLE [PageListMap] ADD CONSTRAINT [PageListOptionsPageListMap] FOREIGN KEY (ListID) REFERENCES [PageListOptions] (RID);
ALTER TABLE [PageListMap] ADD CONSTRAINT [TabsPageListMap] FOREIGN KEY (TabID) REFERENCES [Tabs] (TabID);

/* Foreign keys for Pages */
ALTER TABLE [Pages] ADD CONSTRAINT [CategoriesPages] FOREIGN KEY (Category) REFERENCES [Categories] (CategoryID);

/* Foreign keys for Pictures */
ALTER TABLE [Pictures] ADD CONSTRAINT [PagesPictures] FOREIGN KEY (PageID) REFERENCES [Pages] (RID);

/* Foreign keys for Schedules */
ALTER TABLE [Schedules] ADD CONSTRAINT [PagesSchedules] FOREIGN KEY (PageID) REFERENCES [Pages] (RID);

/* Foreign keys for TabCategoryMap */
ALTER TABLE [TabCategoryMap] ADD CONSTRAINT [CategoriesTabCategoryMap] FOREIGN KEY (CategoryID) REFERENCES [Categories] (CategoryID);
ALTER TABLE [TabCategoryMap] ADD CONSTRAINT [TabsTabCategoryMap] FOREIGN KEY (TabID) REFERENCES [Tabs] (TabID);

/* Views */
DROP VIEW [Tabs Query] /**WEAK**/;
CREATE VIEW [Tabs Query] AS
SELECT Tabs.TabID, Tabs.Name, Tabs.SortName
FROM Tabs;



DROP TABLE [Gallery] /**WEAK**/;
CREATE TABLE [Gallery] (
[RID] COUNTER NOT NULL,
[Title] TEXT(255) NOT NULL,
[EventDate] DATETIME NOT NULL,
[Description] LONGTEXT,
[Parent] INTEGER,
[Disabled] BIT NOT NULL DEFAULT 0,
[CoverPictureID] INTEGER,
[PictureCount] INTEGER,
[SubGalleryCount] INTEGER,
[SubPictureCount] INTEGER,
[DateCreated] DATETIME NOT NULL,
[DateModified] DATETIME NOT NULL
);
CREATE UNIQUE INDEX [PrimaryKey] ON [Gallery] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE UNIQUE INDEX [RID] ON [Gallery] (RID ASC);


DROP TABLE [GalleryCategories] /**WEAK**/;
CREATE TABLE [GalleryCategories] (
[RID] COUNTER NOT NULL,
[Category] TEXT(64) NOT NULL,
[Description] TEXT(255) NOT NULL,
[Disabled] BIT NOT NULL DEFAULT 0
);
CREATE UNIQUE INDEX [PrimaryKey] ON [GalleryCategories] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE UNIQUE INDEX [RID] ON [GalleryCategories] (RID ASC);

DROP TABLE [GalleryCategoryMap] /**WEAK**/;
CREATE TABLE [GalleryCategoryMap] (
[RID] COUNTER NOT NULL,
[CategoryID] INTEGER NOT NULL,
[GalleryID] INTEGER NOT NULL
);
CREATE UNIQUE INDEX [PrimaryKey] ON [GalleryCategoryMap] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE INDEX [GalleryCategoryMapGalleryID] ON [GalleryCategoryMap] (GalleryID ASC);
CREATE INDEX [RID] ON [GalleryCategoryMap] (RID ASC);

DROP TABLE [GalleryPictures] /**WEAK**/;
CREATE TABLE [GalleryPictures] (
[RID] COUNTER NOT NULL,
[GalleryID] INTEGER NOT NULL,
[File] TEXT(40) NOT NULL,
[Caption] TEXT(255),
[Sequence] INTEGER DEFAULT 0,
[Disabled] BIT NOT NULL DEFAULT 0,
[DateAdded] DATETIME NOT NULL
);
CREATE UNIQUE INDEX [PrimaryKey] ON [GalleryPictures] (RID ASC) WITH PRIMARY DISALLOW NULL;
CREATE UNIQUE INDEX [RID] ON [GalleryPictures] (RID ASC);
CREATE INDEX [GalleryID] ON [GalleryPictures] (GalleryID ASC);

DROP TABLE [GalleryShuffle] /**WEAK**/;
CREATE TABLE [GalleryShuffle] (
[RID] COUNTER NOT NULL,
[PictureID] INTEGER NOT NULL
);
CREATE UNIQUE INDEX [PrimaryKey] ON [GalleryShuffle] (RID ASC) WITH PRIMARY DISALLOW NULL;


/* Access SQL export data end. */
