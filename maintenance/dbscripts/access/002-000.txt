

CREATE TABLE PageListOptions
(
	[RID] AUTOINCREMENT NOT NULL,
	[ListName] TEXT(255) WITH COMPRESSION NOT NULL,
	[Details] MEMO WITH COMPRESSION,
	CONSTRAINT [PrimaryKey] PRIMARY KEY( [RID] ),
	CONSTRAINT [RID] UNIQUE ([RID])
);



CREATE TABLE PageListMap
(
	[RID] AUTOINCREMENT  NOT NULL,
	[TabID] LONG  NOT NULL,
	[ListID] LONG  NOT NULL,
	CONSTRAINT [PrimaryKey] PRIMARY KEY( [RID] )
);

CREATE INDEX [PageListMapTabID] ON [PageListMap]([TabID]);
CREATE INDEX [RID] ON [PageListMap]([RID]);

#note: FOREIGN KEYs are always One-to-Many relationships

ALTER TABLE [PageListMap]
	ADD CONSTRAINT [PageListOptionsPageListMap]
	FOREIGN KEY NO INDEX ([ListID]) 
		REFERENCES [PageListOptions] ([RID]);
ALTER TABLE [PageListMap]
	ADD CONSTRAINT [TabsPageListMap]
	FOREIGN KEY NO INDEX ([TabID]) 
		REFERENCES [Tabs] ([TabID]);



INSERT INTO PageListOptions
	( ListName, Details )
	VALUES
	( "Category + Page [ascending]", "IIF(ISNULL(Categories.SortName),'',Categories.SortName)+Categories.Name, IIF(ISNULL(Pages.SortName),'',Pages.SortName)+Pages.PageName" );

INSERT INTO PageListOptions
	( ListName, Details )
	VALUES
	( "Category [asc], Page [desc]", "IIF(ISNULL(Categories.SortName),'',Categories.SortName)+Categories.Name, IIF(ISNULL(Pages.SortName),'',Pages.SortName)+Pages.PageName desc" );

INSERT INTO PageListOptions
	( ListName, Details )
	VALUES
	( "Category + Page [descending]", "IIF(ISNULL(Categories.SortName),'',Categories.SortName)+Categories.Name desc, IIF(ISNULL(Pages.SortName),'',Pages.SortName)+Pages.PageName desc" );

INSERT INTO PageListOptions
	( ListName, Details )
	VALUES
	( "Event Date [simple]", "IIF(ISNULL(Schedules.DateEvent),DATE()+10000,Schedules.DateEvent), IIF(ISNULL(Pages.SortName),'',Pages.SortName)+Pages.PageName" );

INSERT INTO PageListOptions
	( ListName, Details )
	VALUES
	( "Event Date [Dynamic]", 
		"IIF(ISNULL(Schedules.DateEvent),IIF(ISNULL(Schedules.DateBegin),Pages.DateModified,Schedules.DateBegin),IIF(DATE() <= Schedules.DateEvent,IIF(DATEDIFF('d',DATE(),Schedules.DateEvent) < 7,DATEADD('d',8-DATEDIFF('d',DATE(),Schedules.DateEvent),Schedules.DateEvent),DateModified),Pages.DateModified)), IIF(ISNULL(Pages.SortName),'',Pages.SortName)+Pages.PageName" );

INSERT INTO PageListOptions
	( ListName, Details )
	VALUES
	( "Date Modified [desc]", "Pages.DateModified desc, IIF(ISNULL(Pages.SortName),'',Pages.SortName)+Pages.PageName" );


UPDATE Version
	SET Major = 3, Minor = 0, Fixes = 0;
