HansoftExport
=============

About this program
------------------
This program is an Hansoft SDK client program that connects to the Hansoft server and exports a Hansoft report or find query to Excel.  This is similar
to what you can do in the Hansoft client but this utility is primarily useful if you, for example using the Windows Task Scheduler, want to do
the export regularly and then the exported data is then prepared for presentation using, e.g., an Excel Macro. Note that to use this program you need
to have the SDK option enabled on your Hansoft Server and also have an SDK user created in your database.

There is some help displayed if you invoke the program without specifying any parameters. See also below.

This utility exports all active columns without sorting or grouping regardeless of if that is specified in a report given as a parameter.

The columns will be exported in the following order:
* The non-hidable built-in columns. These columns are different depending on what view (Agile, Scheduled, QA, Product Backlog) you are exporting. 
* Active hidable built-in columns
* Active custom columns

Terms and conditions
--------------------
Hansoft Export is licensed under what is known as an MIT License as stated in the [LICENSE.md](LICENSE.md).

This program is not part of the official Hansoft product or subject to its license agreement.
The program is provided as is and there is no obligation on Hansoft AB to provide support, update or enhance this program.

Building the program
--------------------
The program can be built with the freely available [Visual Studio Express 2012 for Desktop] [1]. The export to Excel is done
with [EPPlus] [2] which you will need to download separately and then change the reference
to the EPPlus DLL in the Visual Studio Project as needed. You will also need the [Hansoft SDK] [3] to be able to build the program. You will
also need to update the references to the appropriate 
Hansoft SDK DLL in the Visual Studio project (typically: HPMSdkManaged.x86) and make sure that the Hansoft SDK DLL (typically HPMSdk.x86.dll) is
in the same directory as your executable. Further the program is also based on the [ObjectWrapper] [4] and [SimpleLogging] [5] libraries
which are available on GitHub. You will need to download the source code for these libraries and build them and update the references from
the HansoftExport project to the built libraries.

[1]: http://www.microsoft.com/visualstudio/eng/products/visual-studio-express-for-windows-desktop  "Visual Studio Express 2012 for Desktop"
[2]: http://epplus.codeplex.com/                                                                   "EPPlus"
[3]: http://hansoft.com/support/downloads/                                                         "Hansoft SDK"
[4]: http://github.com/Hansoft/Hansoft-ObjectWrapper                                               "ObjectWrapper"
[5]: http://github.com/Hansoft/Hansoft-SimpleLogging                                               "SimpleLogging"

Usage
-----
	HansoftExport -c<server>:<port>:<database>:<sdk user>:<pwd> -p<project>:[a|s|b|q] -r<report>:<user>|-f<query> -o:<file>

	This utility exports the data of a Hansoft report or a Find query to Excel. All active columns in Hansoft will be
	exported regardless of what columns that has been defined to be visible in the report. There is no guruantueed column
	order but the order will be the same as long as the set of active columns remain unchanged. If any sorting or grouping
	is defined in the report this will also be ignored.

	If any parameter values contain spaces, then the parameter value in question need to be double quouted. Colons are not
	allowed in parameter values.

	Options -c, -p, and -o must always be specified and one of the options -r and-f must also be specified.

	-c Specifies what hansoft database to connect to and the sdk user to be used
	<server>       : IP or DNS name of the Hansoft server
	<port>         : The listen port of the Hansoft server
	<database>     : Name of the Hansoft Database to get data from
	<sdk user>     : Name of the Hansoft SDK User account
	<pwd>          : Password of the Hansoft SDK User account

	-p Specifies the Hansoft project and view to get data from
	<project>      : Name of the Hansoft project
	a              : Get data from the Agile project view
	s              : Get data from the Scheduled project view
	b              : Get data from the Product Backlog
	q              : Get data from the Qaulity Assurance section

	-r Get the data of a Hansoft report
	<report>       : The name of the report
	<user>         : The name of the user that has created the report

	-f Get the data of a Hansoft find query
	<find>         : The query
	Note: if the query expression contains double quoutes, they should be replaced with single quoutes when using this
	      utility.

	-o Specifies the name of the Excel output file
	<file>         : File name

Examples
--------
Find all high priority bugs in the project MyProject in the database My Database where the server is running on the
localmachine on port 50257, save the output to the file Bugs.xslx:

	HansoftExport -clocalhost:50257:"My Database":sdk:sdk -pMyProject:q -f"Bugpriority='High priority'" -oBugs.xlsx

Export all items from the report PBL (defined by Manager Jim) in the product backlog of the project MyProject in the
database My Database found on the server running on the local machine at port 50257:

	HansoftExport -clocalhost:50257:"My Database":sdk:sdk -pMyProject:b -rPBL:"Manager Jim" -oPBL.xlsx

