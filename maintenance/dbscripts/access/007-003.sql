


INSERT INTO [AuthorizeResourceMap]
	( [GroupID], [ResourceID] )
	SELECT [RID], (SELECT [RID] FROM [AuthorizeResource] WHERE [Name] = "Calendar" )
	FROM [AuthorizeGroup]
	WHERE [Name] = "Calendar Editor";



UPDATE [Version]
	SET Major = 7, Minor = 4, Fixes = 0;
