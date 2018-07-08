
# for Version 1.0 we simply add the fixes column to the Version table

ALTER TABLE [Version]
	ADD COLUMN [Fixes] INTEGER NULL;

ALTER TABLE [Version]
	ALTER COLUMN [Minor]  INTEGER NOT NULL;



UPDATE Version
	SET Major = 1, Minor = 0, Fixes = 0;
