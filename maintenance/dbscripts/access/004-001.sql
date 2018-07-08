# add the Gallery Category table

CREATE TABLE GalleryCategories
(
	[RID] AUTOINCREMENT NOT NULL,
	[Category] TEXT(64) WITH COMPRESSION NOT NULL,
	[Description] TEXT(255) WITH COMPRESSION NOT NULL,
	[Disabled] YESNO  DEFAULT 0,
	CONSTRAINT [PrimaryKey] PRIMARY KEY( [RID] ),
	CONSTRAINT [RID] UNIQUE ([RID])
);


CREATE TABLE GalleryCategoryMap
(
	[RID] AUTOINCREMENT  NOT NULL,
	[CategoryID] LONG  NOT NULL,
	[GalleryID] LONG  NOT NULL,
	CONSTRAINT [PrimaryKey] PRIMARY KEY( [RID] )
);

CREATE INDEX [GalleryCategoryMapGalleryID] ON [GalleryCategoryMap]([GalleryID]);
CREATE INDEX [RID] ON [GalleryCategoryMap]([RID]);

#note: FOREIGN KEYs are always One-to-Many relationships

ALTER TABLE [GalleryCategoryMap]
	ADD CONSTRAINT [GalleryGalleryCategoryMap]
	FOREIGN KEY NO INDEX ([GalleryID]) 
		REFERENCES [Gallery] ([RID]);
ALTER TABLE [GalleryCategoryMap]
	ADD CONSTRAINT [GalleryCategoriesGalleryCategoryMap]
	FOREIGN KEY NO INDEX ([CategoryID]) 
		REFERENCES [GalleryCategories] ([RID]);


# now we need to add a new column to the gallery pictures table

ALTER TABLE [GalleryPictures]
	ADD COLUMN [DateAdded] DATETIME  NOT NULL;


# add an optional cover picture to the gallery along with some administrative stuff

ALTER TABLE [Gallery]
	ADD COLUMN [CoverPictureID] LONG  NULL;
ALTER TABLE [Gallery]
	ADD CONSTRAINT [GalleryPicturesGallery]
	FOREIGN KEY NO INDEX ([CoverPictureID]) REFERENCES
		[GalleryPictures] ([RID]);

ALTER TABLE [Gallery]
	ADD COLUMN [PictureCount]  LONG  NULL;

ALTER TABLE [Gallery]
	ADD COLUMN [SubGalleryCount]  LONG  NULL;

ALTER TABLE [Gallery]
	ADD COLUMN [SubPictureCount]  LONG  NULL;

ALTER TABLE [Gallery]
	ADD COLUMN [DateCreated] DATETIME  NOT NULL;

ALTER TABLE [Gallery]
	ADD COLUMN [DateModified] DATETIME  NOT NULL;

# add the constraint that we missed earlier

ALTER TABLE [GalleryShuffle]
	ADD CONSTRAINT [GalleryPicturesGalleryShuffle]
	FOREIGN KEY NO INDEX ([PictureID]) REFERENCES
		[GalleryPictures] ([RID]);



# update to version 5 because we added tables

UPDATE Version
	SET Major = 5, Minor = 0, Fixes = 0;

