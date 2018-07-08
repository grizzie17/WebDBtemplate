
INSERT INTO [AuthorizeResource]
	( [Name], [Abbrev], [Description], [DateModified] )
	VALUES
	( "ChapterMaint", "CHPm", "Maintain Chapter Info Pages", #1/3/2013# );


INSERT INTO [AuthorizeGroup]
	( [Name], [Description], [DateModified] )
	VALUES
	( "Chapter Editor", "Responsible for maintaining Chapter Info pages", #1/3/2013# );


INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "ChapterMaint" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Server Admin";


INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "NewsAuthor" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Server Admin";


INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "ChapterMaint" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Webmaster";


INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "NewsAuthor" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Webmaster";




UPDATE [Version]
	SET Major = 7, Minor = 4, Fixes = 0;

