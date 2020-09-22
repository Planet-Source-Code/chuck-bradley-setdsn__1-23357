ODBC Test Project


This project demonstrates some code concepts such as use of classes,
functions, passing variables, and registry control for setting
SYSTEM DSN's under ODBC.

Please be aware that this code does read/write to the registry
and that care must be taken if the code is changed or used.
Do so at your own risk. I will not be responsible for any
problems resulting from such.

I built this class when doing alot of SQL Server 7
DB work and needed to set ODBC connections on the fly.

It prevented alot of help desk calls in that we did not
have to go out and set up ODBC on each machine.

I wrapped a little project around the class to show how to use it
as well as give someone some quick cut/paste code if they need it.

The class and maybe even the project could provide a foundation
for alot more functionality if someone wants to add to them.

To use the app, set the parms in the text boxes and click on
SetDSN.  This will set a SYSTEM DSN in ODBC.  No validation
for the parms, driver path, or anything else is done.
If the parms are OK, a System DSN will be set. You can
look under Control Panel | ODBC Data Sources to view it.

To use the GetServer command, type in a known DSN in the text box
on the left and click on GetServer. A msgbox will pop up with the
server name on it.  If you have just successfully set a DSN with SetDSN,
you can click on GetServer to look at it the server name you just entered.

VB Rocks.

Regards,
Chuck Bradley.


