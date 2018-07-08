


UPDATE [PageListOptions]
	SET [Details] = "IIF(ISNULL(Pages.Sortname),FORMAT(IIF(ISNULL(Schedules.DateEvent),IIF(ISNULL(Schedules.DateBegin),IIF(DATEDIFF('d',DATE(),Pages.DateCreated) < 7,DATEADD('d', -50*(7-DATEDIFF('d',DATE(),Pages.DateCreated)),Pages.DateCreated),DATEADD('d',3*DATEDIFF('d',Pages.DateCreated,DATE()),Pages.DateCreated)),IIF(DATEDIFF('d',DATE(),Schedules.DateBegin) < 7,DATEADD('d', -50*(7-DATEDIFF('d',DATE(),Schedules.DateBegin)),Schedules.DateBegin),DATEADD('d',3*DATEDIFF('d',Schedules.DateBegin,DATE()),Schedules.DateBegin))),IIF(DATE() <= Schedules.DateEvent,IIF(DATEDIFF('d',DATE(),Schedules.DateEvent) < 7,DATEADD('d', -100*(7-DATEDIFF('d',DATE(),Schedules.DateEvent)),Schedules.DateEvent),IIF(ISNULL(Schedules.DateBegin),Schedules.DateEvent,IIF(DATEDIFF('d',Schedules.DateBegin,DATE()) < 8,DATEADD('d',2*(7-DATEDIFF('d',DATE(),Schedules.DateEvent)),Schedules.DateEvent),Schedules.DateEvent))),Schedules.DateEvent)),'yyyy/mm/dd-hh:nn:ss'),Pages.SortName)&Pages.PageName"
	WHERE [ListName] = "Event Date [Dynamic]";


UPDATE [Version]
	SET Major = 7, Minor = 3, Fixes = 0;
