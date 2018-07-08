SET @OLD_UNIQUE_CHECKS=@@UNIQUE_CHECKS, UNIQUE_CHECKS=0;
SET @OLD_FOREIGN_KEY_CHECKS=@@FOREIGN_KEY_CHECKS, FOREIGN_KEY_CHECKS=0;
SET @OLD_SQL_MODE=@@SQL_MODE, SQL_MODE='TRADITIONAL,ALLOW_INVALID_DATES';


-- SET GLOBAL MAX_ALLOWED_PACKET=(15*1024*1024) ;

-- DROP SCHEMA IF EXISTS `dbtemplate` ;
-- CREATE SCHEMA IF NOT EXISTS `dbtemplate` DEFAULT CHARACTER SET utf8 ;
-- USE `dbtemplate` ;

-- -----------------------------------------------------
-- Table `authorizegroup`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `authorizegroup` ;

CREATE TABLE IF NOT EXISTS `authorizegroup` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `Name` VARCHAR(40) NOT NULL,
  `Description` VARCHAR(255) NULL DEFAULT NULL,
  `DateModified` DATETIME NOT NULL,
  `Disabled` BIT(1) NOT NULL DEFAULT b'0',
  UNIQUE INDEX `RID` (`RID` ASC),
  INDEX `authorizegroupName` (`Name` ASC),
  INDEX `authorizegroupRID` (`RID` ASC))
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- -----------------------------------------------------
-- Table `authorizeusers`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `authorizeusers` ;

CREATE TABLE IF NOT EXISTS `authorizeusers` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `Username` VARCHAR(32) NOT NULL,
  `Password` VARCHAR(32) NOT NULL,
  `Email` VARCHAR(128) NOT NULL,
  `Description` VARCHAR(255) NULL DEFAULT NULL,
  `DateCreated` DATETIME NOT NULL,
  `DateModified` DATETIME NOT NULL,
  `Disabled` BIT(1) NOT NULL DEFAULT b'0',
  `Lastname` VARCHAR(32) NULL DEFAULT NULL,
  `Firstname` VARCHAR(32) NULL DEFAULT NULL,
  UNIQUE INDEX `RID` (`RID` ASC),
  INDEX `authorizeusersRID` (`RID` ASC),
  INDEX `authorizeusersUsername` (`Username` ASC))
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;

INSERT INTO `authorizeusers`
	( `UserName`, `Password`, `Email`, `DateCreated`, `DateModified`, `Lastname`, `Firstname` )
	VALUES
	( "grizzie17", "grizzly17", "john@grizzlyweb.com", NOW(), NOW(), "Griswold", "John" );

INSERT INTO `authorizeusers`
	( `UserName`, `Password`, `Email`, `DateCreated`, `DateModified`, `Lastname`, `Firstname` )
	VALUES
	( "webby17", "grizzly17", "john@griwold.me", NOW(), NOW(), "Griswold", "John" );

INSERT INTO `authorizeusers`
	( `UserName`, `Password`, `Email`, `DateCreated`, `DateModified`, `Lastname`, `Firstname` )
	VALUES
	( "admin123", "123.admin!321", "", NOW(), NOW(), "Admin", "Admin" );



-- -----------------------------------------------------
-- Table `authorizegroupmap`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `authorizegroupmap` ;

CREATE TABLE IF NOT EXISTS `authorizegroupmap` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `UserID` INT(11) NOT NULL,
  `GroupID` INT(11) NOT NULL,
  UNIQUE INDEX `RID` (`RID` ASC),
  INDEX `authorizegroupmapRID` (`RID` ASC),
  INDEX `authorizegroupmapUserID` (`UserID` ASC),
  INDEX `Groupsauthorizegroupmap` (`GroupID` ASC),
  CONSTRAINT `Usersauthorizegroupmap`
    FOREIGN KEY (`UserID`)
    REFERENCES `authorizeusers` (`RID`)
	ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT `Groupsauthorizegroupmap`
    FOREIGN KEY (`GroupID`)
    REFERENCES `authorizegroup` (`RID`)
	ON DELETE CASCADE ON UPDATE CASCADE)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- -----------------------------------------------------
-- Table `authorizelog`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `authorizelog` ;

CREATE TABLE IF NOT EXISTS `authorizelog` (
  `UserID` INT(11) NOT NULL,
  `DateLogon` DATETIME NOT NULL,
  `DateFail` DATETIME NOT NULL,
  `FailCount` INT(11) NOT NULL,
  UNIQUE INDEX `AuthorizeLogUserID` (`UserID` ASC),
  CONSTRAINT `authorizeusersAuthorizeLog`
    FOREIGN KEY (`UserID`)
    REFERENCES `authorizeusers` (`RID`))
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- -----------------------------------------------------
-- Table `authorizeremember`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `authorizeremember` ;

CREATE TABLE IF NOT EXISTS `authorizeremember` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `UserID` INT(11) NOT NULL,
  `Cookie` VARCHAR(40) NOT NULL,
  `DateModified` DATETIME NOT NULL,
  UNIQUE INDEX `RID` (`RID` ASC),
  INDEX `AuthorizeRememberRID` (`RID` ASC),
  INDEX `AuthorizeRememberUserID` (`UserID` ASC),
  CONSTRAINT `UsersAuthorizeRemember`
    FOREIGN KEY (`UserID`)
    REFERENCES `authorizeusers` (`RID`)
	ON DELETE CASCADE ON UPDATE CASCADE)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- -----------------------------------------------------
-- Table `authorizeresource`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `authorizeresource` ;

CREATE TABLE IF NOT EXISTS `authorizeresource` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `Name` VARCHAR(40) NOT NULL,
  `Abbrev` VARCHAR(4) NOT NULL,
  `Description` VARCHAR(255) NULL DEFAULT NULL,
  `DateModified` DATETIME NOT NULL,
  `Disabled` BIT(1) NOT NULL DEFAULT b'0',
  UNIQUE INDEX `RID` (`RID` ASC),
  INDEX `authorizeresourceName` (`Name` ASC),
  INDEX `authorizeresourceRID` (`RID` ASC))
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- -----------------------------------------------------
-- Table `authorizeresourcemap`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `authorizeresourcemap` ;

CREATE TABLE IF NOT EXISTS `authorizeresourcemap` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `GroupID` INT(11) NOT NULL,
  `ResourceID` INT(11) NOT NULL,
  UNIQUE INDEX `RID` (`RID` ASC),
  INDEX `authorizeresourcemapGroupID` (`GroupID` ASC),
  INDEX `authorizeresourcemapRID` (`RID` ASC),
  INDEX `Resourcesauthorizeresourcemap` (`ResourceID` ASC),
  CONSTRAINT `Resourcesauthorizeresourcemap`
    FOREIGN KEY (`ResourceID`)
    REFERENCES `authorizeresource` (`RID`)
	ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT `Groupsauthorizeresourcemap`
    FOREIGN KEY (`GroupID`)
    REFERENCES `authorizegroup` (`RID`)
	ON DELETE CASCADE ON UPDATE CASCADE)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- -----------------------------------------------------
-- Table `categories`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `categories` ;

CREATE TABLE IF NOT EXISTS `categories` (
  `CategoryID` INT(11) NOT NULL AUTO_INCREMENT,
  `Name` VARCHAR(50) NOT NULL,
  `ShortName` VARCHAR(30) NULL DEFAULT NULL,
  `ShortDesc` LONGTEXT NULL DEFAULT NULL,
  `Icon` INT(11) NULL DEFAULT NULL,	-- imageID
  `SortName` TINYTEXT NULL DEFAULT NULL,
  `Revision` INT(11) NOT NULL DEFAULT 0,
  `Disabled` BIT(1) NOT NULL DEFAULT b'0',
  UNIQUE INDEX `CategoryID` (`CategoryID` ASC))
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;

INSERT INTO `categories` ( `Name`, `ShortDesc` ) VALUES ( '@Library', 'Images for global use' );
INSERT INTO `categories` ( `Name` ) VALUES ( '@External' );
INSERT INTO `categories` ( `Name`, `ShortDesc` ) VALUES ( 'Home', 'Special content on home page' );
INSERT INTO `categories` ( `Name`, `ShortDesc` ) VALUES ( 'Calendar', 'Special content on calendar page' );
INSERT INTO `categories` ( `Name`, `ShortDesc` ) VALUES ( 'Announcements', 'Content for Home and Extra! pages' );
INSERT INTO `categories` ( `Name`, `ShortDesc` ) VALUES ( 'Newsletters', 'Content for Newsletters pages' );


-- -----------------------------------------------------
-- Table `config`
-- -----------------------------------------------------

DROP TABLE IF EXISTS `config` ;

CREATE TABLE IF NOT EXISTS `config` (
	`RID` INT(11) NOT NULL AUTO_INCREMENT,
	`Name` VARCHAR(64) NOT NULL,
	`Access` VARCHAR(4) NOT NULL DEFAULT 'SRVe',
	`Description` VARCHAR(128) NOT NULL,
	`Type` INT(11) NOT NULL DEFAULT 0,		-- 0=string, 1=integer
	`Maximum` INT(11) NULL DEFAULT NULL,	-- string length, max numeric value
	`DataValue` VARCHAR(255) NULL DEFAULT NULL,
	UNIQUE INDEX `RID` (`RID` ASC),
	INDEX `configureName` (`Name` ASC)
)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('Domain', 'SRVe', 'Domain Name', 0, NULL, 'Unknown.tld' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('SiteName', 'CFGe', 'Site Name', 0, NULL, 'Unknown Site' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('SiteTitle', 'CFGe', 'Site Title (page header)', 0, NULL, 'Unknown Site' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('ShortSiteName', 'CFGe', 'Site Org. Designation', 0, NULL, 'Chapter XX-?' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('SiteMotto', 'CFGe', 'Site Motto', 0, NULL, 'This is Your Motto!' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('SiteCity', 'CFGe', 'Site City, State', 0, NULL, 'Anywhereville, XX' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('CopyrightStartYear', 'CFGe', 'Site Copyright Starting Year', 1, 9999, '2001' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('MailServer', 'SRVe', 'Mail Server', 0, NULL, 'mail.unknown.tld' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('MailboxUpcomingEvents', 'SRVe', 'Mailbox Upcoming Events', 0, NULL, 'UpcomingEvents@Unknown.tld' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('MailboxNewsletter', 'SRVe', 'Mailbox Newsletter Notification', 0, NULL, 'NewsletterNotice@Unknown.tld' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('AnnonUser', 'SRVe', 'Notification Email', 0, NULL, 'notifications@unknown.tld' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('AnnonPW', 'SRVe', 'Notification Password', 0, NULL, '' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('SiteZip', 'CFGe', 'Site ZIP (Used for Weather)', 0, 5, '35801' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('SiteChapter', 'CFGe', 'Site Chapter', 0, 5, 'AL-X' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('SiteChapterID', 'CFGe', 'Site Chapter ID', 0, 2, 'X' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('SiteTabNewsletters', 'CFGe', 'Tab Newsletters', 0, NULL, 'Newsletters' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('SiteTabAnnouncements', 'CFGe', 'Tab Announcements', 0, NULL, 'Extra!' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('SiteRSSAnnounceInclude', 'CFGe', 'RSS Announcements (Filter Include)', 0, NULL, '' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('SiteRSSAnnounceExclude', 'CFGe', 'RSS Announcements (Filter Exclude)', 0, NULL, 'calendar,personal,plaque' );

INSERT INTO `config` (`Name`, `Access`, `Description`, `Type`, `Maximum`, `DataValue`)
	VALUES ('HomepageLayout', 'CFGe', 'Homepage Layout (a-announce, c-calendar, e-extras)', 0, 3, 'ace' );




-- -----------------------------------------------------
-- Table `gallerypictures`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `gallerypictures` ;

CREATE TABLE IF NOT EXISTS `gallerypictures` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `GalleryID` INT(11) NOT NULL,
  `File` VARCHAR(40) NOT NULL,
  `Caption` VARCHAR(255) NULL DEFAULT NULL,
  `Sequence` INT(11) NULL DEFAULT '0',
  `Disabled` BIT(1) NOT NULL DEFAULT b'0',
  `DateAdded` DATETIME NOT NULL,
  UNIQUE INDEX `RID` (`RID` ASC),
  INDEX `GalleryID` (`GalleryID` ASC),
  CONSTRAINT `GalleryGalleryPictures`
    FOREIGN KEY (`GalleryID`)
    REFERENCES `gallery` (`RID`))
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- -----------------------------------------------------
-- Table `gallery`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `gallery` ;

CREATE TABLE IF NOT EXISTS `gallery` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `Title` VARCHAR(255) NOT NULL,
  `EventDate` DATETIME NOT NULL,
  `Description` LONGTEXT NULL DEFAULT NULL,
  `Parent` INT(11) NULL DEFAULT NULL,
  `Disabled` BIT(1) NOT NULL DEFAULT b'0',
  `CoverPictureID` INT(11) NULL DEFAULT NULL,
  `PictureCount` INT(11) NULL DEFAULT NULL,
  `SubGalleryCount` INT(11) NULL DEFAULT NULL,
  `SubPictureCount` INT(11) NULL DEFAULT NULL,
  `DateCreated` DATETIME NOT NULL,
  `DateModified` DATETIME NOT NULL,
  UNIQUE INDEX `RID` (`RID` ASC),
  INDEX `GalleryPicturesGallery` (`CoverPictureID` ASC),
  CONSTRAINT `GalleryPicturesGallery`
    FOREIGN KEY (`CoverPictureID`)
    REFERENCES `gallerypictures` (`RID`))
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- -----------------------------------------------------
-- Table `gallerycategories`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `gallerycategories` ;

CREATE TABLE IF NOT EXISTS `gallerycategories` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `Category` VARCHAR(64) NOT NULL,
  `Description` VARCHAR(255) NOT NULL,
  `Disabled` BIT(1) NOT NULL DEFAULT b'0',
  UNIQUE INDEX `RID` (`RID` ASC))
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- -----------------------------------------------------
-- Table `gallerycategorymap`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `gallerycategorymap` ;

CREATE TABLE IF NOT EXISTS `gallerycategorymap` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `CategoryID` INT(11) NOT NULL,
  `GalleryID` INT(11) NOT NULL,
  UNIQUE INDEX `RID` (`RID` ASC),
  INDEX `GalleryCategoryMapGalleryID` (`GalleryID` ASC),
  INDEX `GalleryCategoryMapRID` (`RID` ASC),
  INDEX `GalleryCategoriesGalleryCategoryMap` (`CategoryID` ASC),
  CONSTRAINT `GalleryGalleryCategoryMap`
    FOREIGN KEY (`GalleryID`)
    REFERENCES `gallery` (`RID`),
  CONSTRAINT `GalleryCategoriesGalleryCategoryMap`
    FOREIGN KEY (`CategoryID`)
    REFERENCES `gallerycategories` (`RID`))
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- -----------------------------------------------------
-- Table `galleryshuffle`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `galleryshuffle` ;

CREATE TABLE IF NOT EXISTS `galleryshuffle` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `PictureID` INT(11) NOT NULL,
  UNIQUE INDEX `RID` (`RID` ASC),
  INDEX `GalleryPicturesGalleryShuffle` (`PictureID` ASC),
  CONSTRAINT `GalleryPicturesGalleryShuffle`
    FOREIGN KEY (`PictureID`)
    REFERENCES `gallerypictures` (`RID`))
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- -----------------------------------------------------
-- Table `pages`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `pages` ;

CREATE TABLE IF NOT EXISTS `pages` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `PageName` VARCHAR(64) NOT NULL,
  `SortName` VARCHAR(32) NULL DEFAULT NULL,
  `Format` VARCHAR(4) NULL DEFAULT NULL,
  `Title` VARCHAR(128) NOT NULL,
  `Description` VARCHAR(255) NULL DEFAULT NULL,
  `DateCreated` DATETIME NULL DEFAULT NULL,
  `DateModified` DATETIME NULL DEFAULT NULL,
  `Category` INT(11) NULL DEFAULT NULL,
  `Picture` INT(11) NULL DEFAULT NULL,	-- id of image
  `Disabled` BIT(1) NOT NULL DEFAULT b'0',
  `Body` MEDIUMTEXT NULL DEFAULT NULL,
  UNIQUE INDEX `RID` (`RID` ASC),
  INDEX `PagesCategory` (`Category` ASC),
  INDEX `PagesPageName` (`PageName` ASC),
  CONSTRAINT `CategoriesPages`
    FOREIGN KEY (`Category`)
    REFERENCES `categories` (`CategoryID`)
	ON DELETE CASCADE ON UPDATE CASCADE)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


INSERT INTO `pages` ( `PageName`, `Format`, `Title`, `Description`, `Body`, `DateCreated`, `DateModified`, `Category`) 
VALUES ('~Library', 'STFT', 'Library Image Page (Do Not Delete)', '', '', NOW(), NOW(), 
(SELECT `CategoryID` FROM `categories` WHERE `Name` = '@Library')
);

INSERT INTO `pages` ( `PageName`, `Format`, `Title`, `Description`, `Body`, `DateCreated`, `DateModified`, `Category`) 
VALUES ('@External Pages', 'STFT', 'External Pages (Do Not Delete)', '', '', NOW(), NOW(), 
(SELECT `CategoryID` FROM `categories` WHERE `Name` = '@External')
);

INSERT INTO `pages` ( `PageName`, `Format`, `Title`, `Description`, `Body`, `DateCreated`, `DateModified`, `Category`) 
VALUES ('Sample Home Content', 'STFT', 'Sample Home Content', '', 'Sample Home page content', NOW(), NOW(), 
(SELECT `CategoryID` FROM `categories` WHERE `Name` = 'Home')
);

INSERT INTO `pages` ( `PageName`, `Format`, `Title`, `Description`, `Body`, `DateCreated`, `DateModified`, `Category`) 
VALUES ('Sample Calendar Content', 'STFT', 'Sample Calendar Content', '', 'Sample Calendar page content', NOW(), NOW(), 
(SELECT `CategoryID` FROM `categories` WHERE `Name` = 'Calendar')
);

INSERT INTO `pages` ( `PageName`, `Format`, `Title`, `Description`, `Body`, `DateCreated`, `DateModified`, `Category`) 
VALUES ('Sample Announcement', 'STFT', 'Sample Announcement', 'Stuff for the Description', 'Sample Announcement page content', NOW(), NOW(), 
(SELECT `CategoryID` FROM `categories` WHERE `Name` = 'Announcements')
);

INSERT INTO `pages` ( `PageName`, `Format`, `Title`, `Description`, `Body`, `DateCreated`, `DateModified`, `Category`) 
VALUES ('Sample Newsletter', 'STFT', 'Sample Newsletter', NULL, 'Sample Newsletter page content', NOW(), NOW(), 
(SELECT `CategoryID` FROM `categories` WHERE `Name` = 'Newsletters')
);

-- -----------------------------------------------------
-- Table `keywords`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `keywords` ;

CREATE TABLE IF NOT EXISTS `keywords` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `Keyword` VARCHAR(64) NOT NULL,
  `Disabled` BIT(1) NOT NULL DEFAULT b'0',
  UNIQUE INDEX `RID` (`RID` ASC),
  UNIQUE INDEX `KeywordsKeyword` (`Keyword` ASC))
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- -----------------------------------------------------
-- Table `keywordpagemap`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `keywordpagemap` ;

CREATE TABLE IF NOT EXISTS `keywordpagemap` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `KeywordID` INT(11) NOT NULL,
  `PageID` INT(11) NOT NULL,
  UNIQUE INDEX `RID` (`RID` ASC),
  INDEX `KeywordPageMapKeywordID` (`KeywordID` ASC),
  INDEX `KeywordPageMapPageID` (`PageID` ASC),
  CONSTRAINT `PagesKeywordPageMap`
    FOREIGN KEY (`PageID`)
    REFERENCES `pages` (`RID`)
	ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT `KeywordsKeywordPageMap`
    FOREIGN KEY (`KeywordID`)
    REFERENCES `keywords` (`RID`)
	ON DELETE CASCADE ON UPDATE CASCADE)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- -----------------------------------------------------
-- Table `tabs`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `tabs` ;

CREATE TABLE IF NOT EXISTS `tabs` (
  `TabID` INT(11) NOT NULL AUTO_INCREMENT,
  `Name` VARCHAR(50) NULL DEFAULT NULL,
  `SortName` VARCHAR(50) NULL DEFAULT NULL,
  `Description` TEXT NULL DEFAULT NULL,
  `Icon` INT(11) NULL DEFAULT NULL,			-- ImageID
  `Picture` INT(11) NULL DEFAULT NULL,		-- ImageID
  `Background` INT(11) NULL DEFAULT NULL,	-- ImageID
  `Disabled` BIT(1) NOT NULL,
  UNIQUE INDEX `TabID` (`TabID` ASC),
  UNIQUE INDEX `TabName` (`Name` ASC),
  INDEX `TabsSortName` (`SortName` ASC))
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


INSERT INTO `tabs` (`Name`, `SortName`, `Description`, `Icon`, `Picture`, `Background`, `Disabled`) 
VALUES ('Home', '000', '@default.asp', NULL, NULL, NULL, 0);

INSERT INTO `tabs` (`Name`, `SortName`, `Description`, `Icon`, `Picture`, `Background`, `Disabled`) 
VALUES ('Forum', NULL, '@forum.asp', NULL, NULL, NULL, 0);

INSERT INTO `tabs` (`Name`, `SortName`, `Description`, `Icon`, `Picture`, `Background`, `Disabled`) 
VALUES ('Calendar', NULL, '@calendar.asp', NULL, NULL, NULL, 0);

INSERT INTO `tabs` (`Name`, `SortName`, `Description`, `Icon`, `Picture`, `Background`, `Disabled`) 
VALUES ('Extra!', NULL, 'Announcements', NULL, NULL, NULL, 0);

INSERT INTO `tabs` (`Name`, `SortName`, `Description`, `Icon`, `Picture`, `Background`, `Disabled`) 
VALUES ('Newsletters', NULL, 'Newsletters', NULL, NULL, NULL, 0);

-- -----------------------------------------------------
-- Table `pagelistoptions`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `pagelistoptions` ;

CREATE TABLE IF NOT EXISTS `pagelistoptions` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `ListName` VARCHAR(255) NOT NULL,
  `Details` LONGTEXT NULL DEFAULT NULL,
  UNIQUE INDEX `RID` (`RID` ASC))
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;

INSERT INTO pagelistoptions
	( ListName, Details )
	VALUES
	( "Category + Page [ascending]", "CONCAT(IFNULL(categories.SortName,''),categories.Name), CONCAT(IFNULL(pages.SortName,''),pages.PageName)" );

INSERT INTO pagelistoptions
	( ListName, Details )
	VALUES
	( "Category [asc], Page [desc]", "CONCAT(IFNULL(categories.SortName,''),categories.Name), CONCAT(IFNULL(pages.SortName,''),pages.PageName) desc" );

INSERT INTO pagelistoptions
	( ListName, Details )
	VALUES
	( "Category + Page [descending]", "CONCAT(IFNULL(categories.SortName,''),categories.Name) desc, CONCAT(IFNULL(pages.SortName,''),pages.PageName) desc" );

INSERT INTO pagelistoptions
	( ListName, Details )
	VALUES
	( "Event Date [simple]", "IFNULL(schedules.DateEvent,CURDATE()+10000), CONCAT(IFNULL(pages.SortName,''),pages.PageName)" );

INSERT INTO pagelistoptions
	( ListName, Details )
	VALUES
	( "Event Date [Dynamic]", 
		"IFNULL(pages.Sortname,'~'), IF(ISNULL(schedules.DateEvent),DATEDIFF(CURDATE(),pages.DateModified),LEAST(DATEDIFF(schedules.DateEvent,CURDATE()),DATEDIFF(CURDATE(),pages.DateModified))), pages.PageName " );

INSERT INTO pagelistoptions
	( ListName, Details )
	VALUES
	( "Date Modified [desc]", "pages.DateModified desc, CONCAT(IFNULL(Pages.SortName,''),pages.PageName)" );


-- -----------------------------------------------------
-- Table `pagelistmap`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `pagelistmap` ;

CREATE TABLE IF NOT EXISTS `pagelistmap` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `TabID` INT(11) NOT NULL,
  `ListID` INT(11) NOT NULL,
  UNIQUE INDEX `RID` (`RID` ASC),
  INDEX `PageListMapTabID` (`TabID` ASC),
  INDEX `PageListMapRID` (`RID` ASC),
  INDEX `PageListOptionsPageListMap` (`ListID` ASC),
  CONSTRAINT `TabsPageListMap`
    FOREIGN KEY (`TabID`)
    REFERENCES `tabs` (`TabID`)
	ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT `PageListOptionsPageListMap`
    FOREIGN KEY (`ListID`)
    REFERENCES `pagelistoptions` (`RID`)
	ON DELETE CASCADE ON UPDATE CASCADE)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- INSERT INTO `pagelistmap` ( `TabID`, `ListID` )
-- VALUES (
-- 	(SELECT `TabID` FROM `tabs` WHERE `Name`='Extra!'),
-- 	(SELECT `RID` FROM `pagelistoptions` WHERE `ListName`='Event Date [Dynamic]' )
-- );



-- -----------------------------------------------------
-- Table `images`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `images` ;

CREATE TABLE IF NOT EXISTS `images` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `File` VARCHAR(80) NULL DEFAULT NULL,
  `Mime` VARCHAR(40) NULL DEFAULT NULL,
  `DateCreated` DATETIME NOT NULL,
  `Data` MEDIUMBLOB NULL DEFAULT NULL,
  UNIQUE INDEX `imagesRID` (`RID` ASC))
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;




-- -----------------------------------------------------
-- Table `pictures`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `pictures` ;

CREATE TABLE IF NOT EXISTS `pictures` (
  `PictureID` INT(11) NOT NULL AUTO_INCREMENT,
  `Label` VARCHAR(40) NOT NULL,
  `PageID` INT(11) NOT NULL DEFAULT '0',
  `ImageID` INT(11) NOT NULL,
  UNIQUE INDEX `PictureID` (`PictureID` ASC),
  INDEX `PicturesLabel` (`Label` ASC),
  INDEX `PicturesPageID` (`PageID` ASC),
  CONSTRAINT `PagesPictures`
    FOREIGN KEY (`PageID`)
    REFERENCES `pages` (`RID`)
	ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT `ImagesPictures`
    FOREIGN KEY (`ImageID`)
	REFERENCES `images` (`RID`)
	ON DELETE CASCADE ON UPDATE CASCADE
	)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- -----------------------------------------------------
-- Table `schedules`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `schedules` ;

CREATE TABLE IF NOT EXISTS `schedules` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `PageID` INT(11) NOT NULL,
  `DateBegin` DATETIME NULL DEFAULT NULL,
  `DateEnd` DATETIME NULL DEFAULT NULL,
  `DateEvent` DATETIME NULL DEFAULT NULL,
  `Disabled` BIT(1) NOT NULL DEFAULT b'0',
  UNIQUE INDEX `RID` (`RID` ASC),
  INDEX `SchedulesPageID` (`PageID` ASC),
  CONSTRAINT `PagesSchedules`
    FOREIGN KEY (`PageID`)
    REFERENCES `pages` (`RID`)
	ON DELETE CASCADE ON UPDATE CASCADE)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;


-- -----------------------------------------------------
-- Table `tabcategorymap`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `tabcategorymap` ;

CREATE TABLE IF NOT EXISTS `tabcategorymap` (
  `RID` INT(11) NOT NULL AUTO_INCREMENT,
  `TabID` INT(11) NULL DEFAULT '0',
  `CategoryID` INT(11) NULL DEFAULT '0',
  UNIQUE INDEX `RID` (`RID` ASC),
  INDEX `TabCategoryMapCategoryID` (`CategoryID` ASC),
  INDEX `TabCategoryMapTabID` (`TabID` ASC),
  CONSTRAINT `TabsTabCategoryMap`
    FOREIGN KEY (`TabID`)
    REFERENCES `tabs` (`TabID`)
	ON DELETE CASCADE ON UPDATE CASCADE,
  CONSTRAINT `CategoriesTabCategoryMap`
    FOREIGN KEY (`CategoryID`)
    REFERENCES `categories` (`CategoryID`)
	ON DELETE CASCADE ON UPDATE CASCADE)
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;



INSERT INTO `tabcategorymap` (`TabID`, `CategoryID`)
VALUES (
(SELECT `TabID` FROM `tabs` WHERE `Name` = 'Home'),
(SELECT `CategoryID` FROM `categories` WHERE `Name` = '@External')
);

INSERT INTO `tabcategorymap` (`TabID`, `CategoryID`)
VALUES (
(SELECT `TabID` FROM `tabs` WHERE `Name` = 'Home'),
(SELECT `CategoryID` FROM `categories` WHERE `Name` = 'Home')
);

INSERT INTO `tabcategorymap` (`TabID`, `CategoryID`)
VALUES (
(SELECT `TabID` FROM `tabs` WHERE `Name` = 'Calendar'),
(SELECT `CategoryID` FROM `categories` WHERE `Name` = '@External')
);

INSERT INTO `tabcategorymap` (`TabID`, `CategoryID`)
VALUES (
(SELECT `TabID` FROM `tabs` WHERE `Name` = 'Calendar'),
(SELECT `CategoryID` FROM `categories` WHERE `Name` = 'Calendar')
);

INSERT INTO `tabcategorymap` (`TabID`, `CategoryID`)
VALUES (
(SELECT `TabID` FROM `tabs` WHERE `Name` = 'Forum'),
(SELECT `CategoryID` FROM `categories` WHERE `Name` = '@External')
);

INSERT INTO `tabcategorymap` (`TabID`, `CategoryID`)
VALUES (
(SELECT `TabID` FROM `tabs` WHERE `Name` = 'Extra!'),
(SELECT `CategoryID` FROM `categories` WHERE `Name` = 'Announcements')
);

INSERT INTO `tabcategorymap` (`TabID`, `CategoryID`)
VALUES (
(SELECT `TabID` FROM `tabs` WHERE `Name` = 'Newsletters'),
(SELECT `CategoryID` FROM `categories` WHERE `Name` = 'Newsletters')
);



-- -----------------------------------------------------
-- Table `version`
-- -----------------------------------------------------
DROP TABLE IF EXISTS `version` ;

CREATE TABLE IF NOT EXISTS `version` (
  `Major` INT(11) NOT NULL,
  `Minor` INT(11) NOT NULL,
  `Fixes` INT(11) NULL DEFAULT NULL,
  UNIQUE INDEX `IndexVersion` (`Major` ASC))
ENGINE = InnoDB
DEFAULT CHARACTER SET = utf8;

INSERT INTO `version`
	( `Major`, `Minor`, `Fixes` )
	VALUES 
	( 8, 0, 0 );







-- create some resources
--
-- abbrev meanings
--	first three characters are upper case and identify "real" resource
--	last character is lower case and identifies allowed operation
--	"a" = add/edit
--	"e" = edit
--	"m"	= add/edit/delete

INSERT INTO `authorizeresource`
	( `Name`, `Abbrev`, `Description`, `DateModified` )
	VALUES
	( "Resource", "RSCm", "Maintain Resources", NOW() );

INSERT INTO `authorizeresource`
	( `Name`, `Abbrev`, `Description`, `DateModified` )
	VALUES
	( "Users", "USRm", "Maintain Users", NOW() );

INSERT INTO `authorizeresource`
	( `Name`, `Abbrev`, `Description`, `DateModified` )
	VALUES
	( "Passwords", "PWe", "May Change Own Password", NOW() );

INSERT INTO `authorizeresource`
	( `Name`, `Abbrev`, `Description`, `DateModified` )
	VALUES
	( "Server", "SRVe", "Configure Server Settings", NOW() );

INSERT INTO `authorizeresource`
	( `Name`, `Abbrev`, `Description`, `DateModified` )
	VALUES
	( "SiteConfig", "CFGe", "Configure Site Settings", NOW() );

INSERT INTO `authorizeresource`
	( `Name`, `Abbrev`, `Description`, `DateModified` )
	VALUES
	( "Calendar", "CALm", "Calendar Edits", NOW() );

INSERT INTO `authorizeresource`
	( `Name`, `Abbrev`, `Description`, `DateModified` )
	VALUES
	( "Navigation", "NAVm", "Maintain Navigation Structure", NOW() );

INSERT INTO `authorizeresource`
	( `Name`, `Abbrev`, `Description`, `DateModified` )
	VALUES
	( "PageAuthor", "PAGm", "Add and Maintain Pages", NOW() );

INSERT INTO `authorizeresource`
	( `Name`, `Abbrev`, `Description`, `DateModified` )
	VALUES
	( "PageEditor", "PAGe", "Edit Page Content", NOW() );

INSERT INTO `authorizeresource`
	( `Name`, `Abbrev`, `Description`, `DateModified` )
	VALUES
	( "NewsAuthor", "NWSm", "Add and Maintain Newsletter Pages", NOW() );

INSERT INTO `authorizeresource`
	( `Name`, `Abbrev`, `Description`, `DateModified` )
	VALUES
	( "ChapterMaint", "CHPm", "Maintain Chapter Info Pages", NOW() );

INSERT INTO `authorizeresource`
	( `Name`, `Abbrev`, `Description`, `DateModified` )
	VALUES
	( "StyleMaint", "STYm", "Maintain and Edit Site Styles", NOW() );


-- set up some groups

INSERT INTO `authorizegroup`
	( `Name`, `Description`, `DateModified` )
	VALUES
	( "Server Admin", "Server Administrator with lots of special privs", NOW() );

INSERT INTO `authorizegroup`
	( `Name`, `Description`, `DateModified` )
	VALUES
	( "Webmaster", "Webmaster responsible for most of the maintenance of the web content", NOW() );

INSERT INTO `authorizegroup`
	( `Name`, `Description`, `DateModified` )
	VALUES
	( "Newsletter Editor", "Responsible for adding content to the newsletter", NOW() );

INSERT INTO `authorizegroup`
	( `Name`, `Description`, `DateModified` )
	VALUES
	( "Site Editor", "Responsible for checking content on the website", NOW() );

INSERT INTO `authorizegroup`
	( `Name`, `Description`, `DateModified` )
	VALUES
	( "Calendar Editor", "Responsible for Calendar Edits", NOW() );

INSERT INTO `authorizegroup`
	( `Name`, `Description`, `DateModified` )
	VALUES
	( "Chapter Editor", "Responsible for maintaining Chapter Info pages", NOW() );


-- map the resources into the groups

-- Server Admin

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "Resource" )
	FROM `authorizegroup`
	WHERE `Name` = "Server Admin";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "Users" )
	FROM `authorizegroup`
	WHERE `Name` = "Server Admin";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "Passwords" )
	FROM `authorizegroup`
	WHERE `Name` = "Server Admin";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "Server" )
	FROM `authorizegroup`
	WHERE `Name` = "Server Admin";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "SiteConfig" )
	FROM `authorizegroup`
	WHERE `Name` = "Server Admin";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "Calendar" )
	FROM `authorizegroup`
	WHERE `Name` = "Server Admin";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "Navigation" )
	FROM `authorizegroup`
	WHERE `Name` = "Server Admin";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "PageAuthor" )
	FROM `authorizegroup`
	WHERE `Name` = "Server Admin";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "ChapterMaint" )
	FROM `authorizegroup`
	WHERE `Name` = "Server Admin";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "StyleMaint" )
	FROM `authorizegroup`
	WHERE `Name` = "Server Admin";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "NewsAuthor" )
	FROM `authorizegroup`
	WHERE `Name` = "Server Admin";


-- Webmaster

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "Passwords" )
	FROM `authorizegroup`
	WHERE `Name` = "Webmaster";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "SiteConfig" )
	FROM `authorizegroup`
	WHERE `Name` = "Webmaster";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "Calendar" )
	FROM `authorizegroup`
	WHERE `Name` = "Webmaster";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "Navigation" )
	FROM `authorizegroup`
	WHERE `Name` = "Webmaster";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "PageAuthor" )
	FROM `authorizegroup`
	WHERE `Name` = "Webmaster";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "ChapterMaint" )
	FROM `authorizegroup`
	WHERE `Name` = "Webmaster";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "NewsAuthor" )
	FROM `authorizegroup`
	WHERE `Name` = "Webmaster";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "StyleMaint" )
	FROM `authorizegroup`
	WHERE `Name` = "Webmaster";



-- Newsletter Editor

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "Passwords" )
	FROM `authorizegroup`
	WHERE `Name` = "Newsletter Editor";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "NewsAuthor" )
	FROM `authorizegroup`
	WHERE `Name` = "Newsletter Editor";



-- Site Editor

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "Passwords" )
	FROM `authorizegroup`
	WHERE `Name` = "Site Editor";

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "PageEditor" )
	FROM `authorizegroup`
	WHERE `Name` = "Site Editor";


-- Calendar Editor

INSERT INTO `authorizeresourcemap`
	( `GroupID`, `ResourceID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizeresource` WHERE `Name` = "Calendar" )
	FROM `authorizegroup`
	WHERE `Name` = "Calendar Editor";



-- now let's map the groups to a user

INSERT INTO `authorizegroupmap`
	( `UserID`, `GroupID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizegroup` WHERE `Name` = "Server Admin")
	FROM `authorizeusers`
	WHERE `Username` = "grizzie17";

INSERT INTO `authorizegroupmap`
	( `UserID`, `GroupID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizegroup` WHERE `Name` = "Webmaster")
	FROM `authorizeusers`
	WHERE `Username` = "grizzie17";

INSERT INTO `authorizegroupmap`
	( `UserID`, `GroupID` )
	SELECT `RID`, (SELECT `RID` FROM `authorizegroup` WHERE `Name` = "Webmaster")
	FROM `authorizeusers`
	WHERE `Username` = "webby17";












SET SQL_MODE=@OLD_SQL_MODE;
SET FOREIGN_KEY_CHECKS=@OLD_FOREIGN_KEY_CHECKS;
SET UNIQUE_CHECKS=@OLD_UNIQUE_CHECKS;
