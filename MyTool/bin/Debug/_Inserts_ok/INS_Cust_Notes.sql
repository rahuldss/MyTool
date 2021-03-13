Set Identity_insert Cust_Notes On
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(3,2,'kjbgbkmvbkhbvhhvhjvjhv b.k','testcustomer','May 26 2009  6:14:35:510PM','Jan  1 1900 12:00:00:000AM',NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(4,2,'kjbgbkmvbkhbvhhvhjvjhv b.k','testcustomer','May 26 2009  6:14:35:713PM','Jan  1 1900 12:00:00:000AM',NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(18,202,'Prefer not to use Rachel Farrell','emily','Nov 10 2009 12:52:16:000PM',NULL,NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(19,139,'3 Washington Street (Keene)','emily','Nov 17 2009  4:23:15:000PM',NULL,NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(20,24,'2 Industrial Park Dr, Building 2 is the physical location','emily','Jan  6 2010  3:15:25:000PM',NULL,NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(15,18,'test Notes','Admin','Jun 26 2009  3:40:35:710AM','Jan  1 1900 12:00:00:000AM',NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(17,29,'Fax number is for the yellow pod','dawn','Jul  1 2009  1:53:16:000PM',NULL,NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(21,133,'Physical address is 32 Clinton Street.','emily','Jan 25 2010  2:23:40:000PM',NULL,NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(22,548,'Physical address is 101 Merrimack St.','emily','Jan 25 2010  3:23:24:000PM',NULL,NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(23,512,'(Changes in RMS Reporting)

a)	We would like all of the customer classes to be listed on the Filled Referrals report even if there are no requests for that particular class for a given time period.  

b)	We would like the colors within the pie chart to remain consistent for the Filled Referrals report.  For example, we would like Businesses to always appear as green, colleges to appear as blue, medical to appear as red and so on.  All the classes should be assigned a color and that color should not change.

c)	Any report or list should tell you the total number of pages in that report or list.  It is not sufficient for the database to say “1+” when there are multiple pages.

d)	None of the reports should break the page in the middle of a request.  Please add page breaks in so that the whole request appears on one page.

e)	Please see the three attachments “Filled Referrals for January 2010.jpg”, “why not show up.jpg” and “NAFs.jpg” as an illustration that some Cancelled appointments and zero Not Able to Fill requests are not being reported on the Filled Referrals report.  We need accurate numbers for our statistics and currently the database is not pulling accurate numbers.

f)	The file “not CART providers.jpg” shows two interpreters on the list of CART providers.  The toggle switch for CART providers is off for both of the interpreters so why do these interpreters appear on the CART provider list?

g)	I can’t get any results when I try to pull up a report for New/Notified Referrals Cust Class.  Is this report pulling the right data?


a)	We need some reports to look like the report you pull up when you choose “Filled Referrals” - i.e., with a pie chart and not in textform.  Please see the attached examples (“example of textform.jpg” and “example of pie chart report.jpg”) if you need further clarification.  Please also see the Excel spreadsheet “Chart for Reports menu” for a detailed list of those reports and their formats.

b)	The menu item name in the Reports menu needs to be the same name (or title) that appears on the actual report.  For example, the menu item “Service Hours” shows up as Interpreting Service Hours on the actual report.  This happens with Open Referral List too.  The report shows up as Open Referrals, but the menu item is called Open Referral List.  Please see the Excel spreadsheet for clarification on the Menu and Report names.  Bottom line is that they need to be the same name.

c)	We need a list of the Inactive and Not NH Licensed Vendors on the Reports menu.  Currently, there are checkboxes in the vendor’s profile to mark someone as Inactive or Not NH Licensed, but we have no way of pulling up a list to see the individuals who have either or both of those labels.

d)	We would really like the database to record the date the request was assigned by a user.  This would help us because some of the reports rely on the date assigned to pull accurate numbers.  For example, Time Taken to Assign Vendor needs to measure the time from the date the request was added to the database to the date the request was assigned in the database.  Perhaps the database internally keeps track of the date assigned, but it would be nice to see a visual of this.  Perhaps a button could be added to the Billing Information, Date/Time/Location or Vendor tab that lists the date the request was assigned.

e)	Interpreter List By Services: The report should be a piechart of the different services (Interpreter, CART, legal, etc) that are used within a given time frame.  For example, it should show us that 75% of our requests in January 2010 were interpreter requests, 10% were CART requests, 10% were legal requests and the remaining 5% were low-vision requests.  However, it would be helpful to pull up a list (not a report) of all of the vendors classified as interpreters or CART providers or legal interpreters.  So we would like to see a report called Vendor Services by Type (replacing the name Interpreter List By Services) and a list called Vendor Services by Type (List).

f)	Time Taken To Assign Vendor: We need this report to be a piechart showing the following kind of information: 35% of requests in January took 5 days to fill, 10% of requests in January took 3 days to fill, 40% of requests took 10 days to fill.  

g)	Service Hours: We need this report to tell us the number of service hours for a given class (business, medical, college, etc.) within a certain time frame.  We would like a piechart to show the percentages according to class for that time frame.  In other words, we need to know that medical made up 50% of the total service hours for the month of January and that businesses made up 25% of the total service hours and that colleges made up the remaining 25% of service hours.

h)	Referrals by Location: We need this report to be in piechart form and to show, for example, that 75% of assigned referrals in January were in Manchester and 10% of assigned referrals were in Concord and 15% of assigned referrals in January were in Nashua.

i)	Filled Referrals: The totals in Filled Referrals needs to be directly related to the number of filled requests and not the number of interpreters who filled those requests.  

j)   Part B Consumer List: At the end of the Part B Consumer Report we would like to know the numbers of:

Total number of ‘Active’ Part B Consumers
Total number of Part B Consumers listed as ‘Inactive’
Total number of ‘New’ Part B Consumers 

','Admin','Feb 10 2010 12:13:13:000PM','Feb 10 2010 12:13:21:000PM',NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(24,19,'Physical address: 330 Borthwick Ave, Portsmouth, NH 03801','emily','Feb 10 2010  3:35:39:000PM',NULL,NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(25,470,'Somersworth location: 330 Route 108 North, Somersworth, NH 03787','emily','Feb 19 2010  3:56:39:000PM',NULL,NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(26,302,'Nicole - 624-2041, fax 624-2113','emily','Mar  2 2010  1:39:05:000PM',NULL,NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(27,578,'When confirming please CC: to Kathy Yell, kathyy@ywcanh.org as well as confirming to Shirley.','emily','Mar  5 2010 12:09:28:000PM',NULL,NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(28,690,'Physical address is 10 Route 125, Brentwood, NH 03833','emily','Apr 26 2010  4:31:39:000PM',NULL,NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(29,446,'Temporarily housed in Nashua due to construction, normally Hills North is located in Manchester','emily','Jun 17 2010 11:02:36:000AM',NULL,NULL)
INSERT INTO [Cust_Notes] ([NOTEID],[CUSTNMBR],[NOTEDESC],[USERID],[NOTEDATE],[MODFDATE],[IsDeleted])VALUES(10,501,'newtest notes g','Jones','Jun  1 2009 12:12:30:997AM','Jun  1 2009 12:12:54:327AM',NULL)
Set Identity_insert Cust_Notes Off
