# get rid of the old products tables

DROP TABLE [Brands];

DROP TABLE [Gallery];

DROP TABLE [Orders];

DROP TABLE [Specials];

DROP TABLE [Products];



UPDATE Version
	SET Major = 1, Minor = 1, Fixes = 0;
