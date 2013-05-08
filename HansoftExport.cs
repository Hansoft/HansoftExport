using HPMSdk;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;


using Hansoft.SdkUtils;

using HPMInt8 = System.SByte;
using HPMUInt8 = System.Byte;
using HPMInt16 = System.Int16;
using HPMUInt16 = System.UInt16;
using HPMInt32 = System.Int32;
using HPMUInt32 = System.UInt32;
using HPMInt64 = System.Int64;
using HPMUInt64 = System.UInt64;
using HPMError = System.Int32;
using HPMString = System.String;

namespace Hansoft.HsExp
{
    class HsExp
    {
        enum SearchSpec { Report, FindQuery }

        static string sdkUser;
        static string sdkUserPwd;
        static string server;
        static int portNumber;
        static string databaseName;
        static string projectName;
        static string reportUserName;
        static string reportName;
        static string findQuery;
        static EHPMReportViewType viewType;
        static string outputFileName;
        static SearchSpec searchSpec;

        static string usage =
@"Usage:
hsexp -c<server>:<port>:<database>:<sdk user>:<sdk password> -p<project>:[a|s|b|q] -r<report>:<user>|-f<query> -o:<file>

This utility exports the data of a Hansoft report or a Find query to Excel. All active columns in Hansoft will be exported
regardless of what columns that has been defined to be visible in the report. There is no guruantueed column order but the
order will be the same as long as the set of active columns remain unchanged. If any sorting or grouping is defined in the
report this will also be ignored.

If any parameter values contain spaces, then the parameter value in question need to be double quouted. Colons are not
allowed in parameter values.

Options -c, -p, and -o must always be specified and one of the options -r and-f must also be specified.

-c Specifies what hansoft database to connect to and the sdk user to be used
<server>       : IP or DNS name of the Hansoft server
<port>         : The listen port of the Hansoft server
<database>     : Name of the Hansoft Database to get data from
<sdk user>     : Name of the Hansoft SDK User account
<sdk password> : Password of the Hansoft SDK User account

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
Note: if the query expression contains double quoutes, they should be replaced with single quoutes when using this utility.

-o Specifies the name of the Excel output file 
<file>         : File name


Examples:
Find all high priority bugs in the project MyProject in the database My Database where the server is running on the localmachine
on port 50257, save the output to the file Bugs.xslx:

hsexp -clocalhost:50257:""My Database"":sdk:sdk -pMyProject:q -f""Bugpriority='High priority'"" -oBugs.xlsx

Export all items from the report PBL (defined by Manager Jim) in the product backlog of the project MyProject in the database
My Database found on the server running on the local machine at port 50257:

hsexp -clocalhost:50257:""My Database"":sdk:sdk -pMyProject:b -rPBL:""Manager Jim"" -oPBL.xlsx


";

        static void Main(string[] args)
        {
            try
            {
                if (ParseArguments(args))
                {
                    SessionManager.Initialize(sdkUser, sdkUserPwd, server, portNumber, databaseName);
                    SessionManager.Instance.Connect();
                    if (SessionManager.Instance.Connected)
                    {
                        try
                        {
                            DoExport();
                        }
                        catch (Exception e)
                        {
                            Logger.Exception(e);
                        }
                        finally
                        {
                            SessionManager.Instance.CloseSession();
                        }
                    }
                    else
                    {
                        Logger.Warning("Could not open connection to Hansoft");
                    }
                }
                else
                    Logger.DisplayMessage(usage);
            }
            catch (Exception e)
            {
                Logger.Exception(e);
            }
        }

        static void DoExport()
        {            
            HPMUniqueID projId = HPMUtilities.FindProject(projectName);
            if (viewType == EHPMReportViewType.AgileBacklog)
                projId = SessionManager.Instance.Session.ProjectUtilGetBacklog(projId);
            else if (viewType == EHPMReportViewType.AllBugsInProject)
                projId = SessionManager.Instance.Session.ProjectUtilGetQA(projId);
            if (projId != null)
            {
                HPMString findString;
                if (searchSpec == SearchSpec.Report)
                {
                    HPMUniqueID reportUserId = HPMUtilities.FindUser(reportUserName);
                    if (reportUserId == null)
                        throw new ArgumentException("Could not find the user " + reportUserName);
                    HPMReport report = HPMUtilities.FindReport(projId, reportUserId, reportName);
                    if (report == null)
                        throw new ArgumentException("Could not find the report " + reportName + " for user " + reportUserName + " in project " + projectName);
                    findString = SessionManager.Instance.Session.UtilConvertReportToFindString(report, projId, viewType);
                }
                else
                {
                    findString = findQuery.Replace('\'', '"');
                }
                HPMFindContext findContext = new HPMFindContext();
                HPMFindContextData data = SessionManager.Instance.Session.UtilPrepareFindContext(findString, projId, viewType, findContext);
                HPMTaskEnum items = SessionManager.Instance.Session.TaskFind(data, EHPMTaskFindFlag.None);
                ExcelWriter excelWriter = new ExcelWriter();
                if (items.m_Tasks.Length > 0)
                {
                           EHPMProjectGetDefaultActivatedNonHidableColumnsFlag flag;
                    if (viewType == EHPMReportViewType.ScheduleMainProject)
                        flag = EHPMProjectGetDefaultActivatedNonHidableColumnsFlag.ScheduledMode;
                    else if (viewType == EHPMReportViewType.AgileMainProject)
                        flag = EHPMProjectGetDefaultActivatedNonHidableColumnsFlag.AgileMode;
                    else
                        flag = EHPMProjectGetDefaultActivatedNonHidableColumnsFlag.None;


                    EHPMProjectDefaultColumn[] nonHidableColumns = SessionManager.Instance.Session.ProjectGetDefaultActivatedNonHidableColumns(projId, flag).m_Columns;
                    // Need to add the status field for the schedule/agile view and also for the backlog project.
                    if (viewType == EHPMReportViewType.ScheduleMainProject || viewType == EHPMReportViewType.AgileMainProject || viewType == EHPMReportViewType.AgileBacklog)
                    {
                        Array.Resize(ref nonHidableColumns, nonHidableColumns.Length + 1);
                        nonHidableColumns[nonHidableColumns.Length - 1] = EHPMProjectDefaultColumn.ItemStatus;
                    }
                    EHPMProjectDefaultColumn[] activeBuiltinColumns = SessionManager.Instance.Session.ProjectGetDefaultActivatedColumns(projId).m_Columns;
                    HPMProjectCustomColumnsColumn[] activeCustomColumns = SessionManager.Instance.Session.ProjectCustomColumnsGet(projId).m_ShowingColumns;
                      
                    ExcelWriter.Row row = excelWriter.AddRow();

                    foreach (EHPMProjectDefaultColumn builtinCol in nonHidableColumns)
                        row.AddCell(HPMUtilities.GetColumnName(builtinCol));
                    foreach (EHPMProjectDefaultColumn builtinCol in activeBuiltinColumns)
                        row.AddCell(HPMUtilities.GetColumnName(builtinCol));
                    foreach (HPMProjectCustomColumnsColumn customColumn in activeCustomColumns)
                        row.AddCell(HPMUtilities.GetColumnName(projId, customColumn));

                    HPMColumnTextOptions options = new HPMColumnTextOptions();
                    options.m_bForDisplay = true;
                    foreach (HPMUniqueID item in items.m_Tasks)
                    {
                        string description = SessionManager.Instance.Session.TaskGetDescription(item);
                        HPMUniqueID itemRef;
                        if (viewType == EHPMReportViewType.ScheduleMainProject || viewType == EHPMReportViewType.AgileMainProject)
                        {
                            itemRef = SessionManager.Instance.Session.TaskGetProxy(item);
                            if (!itemRef.IsValid())
                                itemRef = SessionManager.Instance.Session.TaskGetMainReference(item);
                        }
                        else
                        {
                            itemRef = SessionManager.Instance.Session.TaskGetMainReference(item);
                        }

                        row = excelWriter.AddRow();
                        foreach (EHPMProjectDefaultColumn builtinCol in nonHidableColumns)
                        {
                            HPMColumn column = new HPMColumn();
                            column.m_ColumnType = EHPMColumnType.DefaultColumn;
                            column.m_ColumnID = (uint)builtinCol;
                            row.AddCell(SessionManager.Instance.Session.TaskRefGetColumnText(itemRef, column, options));
                        }
                        foreach (EHPMProjectDefaultColumn builtinCol in activeBuiltinColumns)
                        {
                            HPMColumn column = new HPMColumn();
                            column.m_ColumnType = EHPMColumnType.DefaultColumn;
                            column.m_ColumnID = (uint)builtinCol;
                            row.AddCell(SessionManager.Instance.Session.TaskRefGetColumnText(itemRef, column, options));
                        }
                        foreach (HPMProjectCustomColumnsColumn customColumn in activeCustomColumns)
                        {
                            string displayString;
                            string dbValue = SessionManager.Instance.Session.TaskGetCustomColumnData(item, SessionManager.Instance.Session.UtilGetColumnHash(customColumn));
                            if (customColumn.m_Type == EHPMProjectCustomColumnsColumnType.DateTime || customColumn.m_Type == EHPMProjectCustomColumnsColumnType.DateTimeWithTime)
                            {
                                ulong ticksSince1970 = SessionManager.Instance.Session.UtilDecodeCustomColumnDateTimeValue(dbValue)*10;
                                DateTime dateTime = new DateTime(1970, 1, 1, 0, 0, 0, DateTimeKind.Utc).AddTicks((long)ticksSince1970);
                                if (customColumn.m_Type == EHPMProjectCustomColumnsColumnType.DateTime)
                                    displayString = dateTime.ToShortDateString();
                                else
                                    displayString = dateTime.ToString();
                            }
                            else if (customColumn.m_Type == EHPMProjectCustomColumnsColumnType.Resources)
                            {
                                displayString ="";
                                HPMResourceDefinitionList resourceList = SessionManager.Instance.Session.UtilDecodeCustomColumnResourcesValue(dbValue);
                                for (int i = 0; i < resourceList.m_Resources.Length; i += 1)
                                {
                                    HPMResourceDefinition resourceDefinition = resourceList.m_Resources[i];
                                    switch (resourceDefinition.m_GroupingType)
                                    {
                                        case EHPMResourceGroupingType.AllProjectMembers:
                                            displayString += "All Project Members";
                                            break;
                                        case EHPMResourceGroupingType.Resource:
                                            displayString += HPMUtilities.GetUserName(resourceDefinition.m_ID);
                                            break;
                                        case EHPMResourceGroupingType.ResourceGroup:
                                            displayString += HPMUtilities.GetGroupName(resourceDefinition.m_ID);
                                            break;
                                    }
                                    if (i < resourceList.m_Resources.Length - 1)
                                        displayString += ", ";
                                }
                            }
                            else if (customColumn.m_Type == EHPMProjectCustomColumnsColumnType.DropList)
                            {
                                displayString = HPMUtilities.DecodeDroplistValue(dbValue, customColumn.m_DropListItems);
                            }
                            else if (customColumn.m_Type == EHPMProjectCustomColumnsColumnType.MultiSelectionDropList)
                            {
                                displayString = "";
                                string[] dbValues = dbValue.Split(new char[] { ';' });
                                for (int i = 0; i < dbValues.Length; i += 1)
                                {
                                    displayString += HPMUtilities.DecodeDroplistValue(dbValues[i], customColumn.m_DropListItems);
                                    if (i < dbValues.Length - 1)
                                        displayString += ", ";
                                }
                            }
                            else
                                displayString = dbValue;

                            row.AddCell(displayString);
                        }
                    }
                }
                excelWriter.SaveAsOfficeOpenXml(outputFileName);
            }
            else
                throw new ArgumentException("Could not find the project " + projectName);
        }


        static bool ParseArguments(string[] args)
        {
//            Console.WriteLine("** Hit return to continue");
//            Console.ReadLine();
            bool cOptionFound = false;
            bool pOptionFound = false;
            bool rOptionFound = false;
            bool fOptionFound = false;
            bool oOptionFound = false;
            foreach (string optionString in args)
            {
                string option = optionString.Substring(0, 2);
                string[] pars = optionString.Substring(2).Split(new char[] { ':' });
                for (int i = 0; i < pars.Length; i += 1)
                {
                    if (pars[i].StartsWith("\"") && pars[i].EndsWith("\""))
                        pars[i] = pars[i].Substring(1, pars[i].Length - 2);
                }
                switch (option)
                {
                    case "-c":
                        if (cOptionFound)
                            throw new ArgumentException("The -c option can only be specified once");
                        if (pars.Length != 5)
                            throw new ArgumentException("The -c option was not specified correctly");
                        cOptionFound = true;
                        server = pars[0];
                        portNumber = Int32.Parse(pars[1]);
                        databaseName = pars[2];
                        sdkUser = pars[3];
                        sdkUserPwd = pars[4];
                        break;
                    case "-p":
                        if (pOptionFound)
                            throw new ArgumentException("The -p option can only be specified once");
                        if (pars.Length != 2)
                            throw new ArgumentException("The -p option was not specified correctly");
                        pOptionFound = true;
                        projectName = pars[0];
                        switch (pars[1])
                        {
                            case "a":
                                viewType = EHPMReportViewType.AgileMainProject;
                                break;
                            case "s":
                                viewType = EHPMReportViewType.ScheduleMainProject;
                                break;
                            case "b":
                                viewType = EHPMReportViewType.AgileBacklog;
                                break;
                            case "q":
                                viewType = EHPMReportViewType.AllBugsInProject;
                                break;
                            default:
                                throw new ArgumentException("An unsupported view type was specified, valid values are one of [a|s|b|q]");
                        }
                        break;
                    case "-r":
                        if (rOptionFound || fOptionFound)
                            throw new ArgumentException("Either the -r option or the -f option can only be specified once");
                        if (pars.Length != 2)
                            throw new ArgumentException("The -r option was not specified correctly");
                        reportName = pars[0];
                        reportUserName = pars[1];
                        searchSpec = SearchSpec.Report;
                        rOptionFound = true;
                        break;
                    case "-f":
                        if (rOptionFound || fOptionFound)
                            throw new ArgumentException("Either the -r option or the -f option can only be specified once");
                        if (pars.Length != 1)
                            throw new ArgumentException("The -f option was not specified correctly");
                        findQuery = pars[0];
                        searchSpec = SearchSpec.FindQuery;
                        fOptionFound = true;
                        break;
                    case "-o":
                        if (oOptionFound)
                            throw new ArgumentException("The -o option can only be specified once");
                        if (pars.Length != 1)
                            throw new ArgumentException("The -o option was not specified correctly");
                        outputFileName = pars[0];
                        oOptionFound = true;
                        break;
                    case "-h":
                        if (oOptionFound)
                            throw new ArgumentException("The -o option can only be specified once");
                        if (pars.Length != 1)
                            throw new ArgumentException("The -o option was not specified correctly");
                        outputFileName = pars[0];
                        oOptionFound = true;
                        break;
                    default:
                        throw new ArgumentException("An unsupported option was specifed: " + option);
                }
            }
            return (cOptionFound && pOptionFound && oOptionFound && (rOptionFound || fOptionFound));
        }
    }
}
