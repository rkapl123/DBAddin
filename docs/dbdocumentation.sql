USE [pubs]
GO
SET ANSI_NULLS ON
GO
SET QUOTED_IDENTIFIER ON
GO
CREATE TABLE [dbo].[dbdocumentation](
	[databasename] [varchar](10) NULL,
	[objecttype] [char](1) NULL,
	[objectname] [varchar](30) NOT NULL,
	[documentation] [varchar](500) NULL,
	[parenttable] [varchar](30) NULL
) ON [PRIMARY]
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (N'pubs', N'T', N'authors', N'table authors contains the book authors', NULL)
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'au_id', N'id of author', N'authors')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'au_lname', N'lastname of author', N'authors')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'au_fname', N'firstname of authors', N'authors')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (N'pubs', N'T', N'discounts', N'discounts per store', NULL)
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'discounttype', N'type of discount', N'discounts')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (N'pubs', N'T', N'employee', N'employees table', NULL)
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (N'pubs', N'T', N'jobs', N'jobs table', NULL)
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (N'pubs', N'T', N'publishers', N'publisher table', NULL)
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (N'pubs', N'T', N'roysched', N'royalty schedules table', NULL)
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (N'pubs', N'T', N'sales', N'sales table', NULL)
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (N'pubs', N'T', N'stores', N'stores table', NULL)
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (N'pubs', N'T', N'titleauthor', N'table linking titles with authors, including the author order and their royalty per title', NULL)
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (N'pubs', N'T', N'titles', N'book titles', NULL)
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (N'pubs', N'T', N'pub_info', N'table with additional information on publishers (logo etc)', NULL)
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'phone', N'phone of author', N'authors')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'city', N'city of author', N'authors')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'state', N'state of author', N'authors')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'zip', N'zip code of author', N'authors')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'contract', N'flag for contract', N'authors')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'discount', N'amount of discount', N'discounts')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'stor_id', N'reference to store', N'discounts')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'emp_id', N'employee id', N'employee')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'fname', N'firstname of employee', N'employee')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'minit', N'middle initial of employee', N'employee')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'lname', N'lastname of employee', N'employee')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'job_id', N'referece to jobs table', N'employee')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'job_lvl', N'level of employee job', N'employee')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'pub_id', N'referece to publishers table', N'employee')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'hire_date', N'hire date of employee', N'employee')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (NULL, N'F', N'job_id', NULL, N'jobs')
GO
INSERT [dbo].[dbdocumentation] ([databasename], [objecttype], [objectname], [documentation], [parenttable]) VALUES (N'pubs', N'T', N'titleview', N'view combining titles, authors and royalty per title (from titleauthor)', NULL)
GO
