using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using System.Xml;

using HPMSdk;
using Hansoft.ObjectWrapper;
using Hansoft.SimpleLogging;

using Hansoft.Excel;

namespace Hansoft.HansoftExport
{
    class HsExp
    {
        enum SearchSpec { Report, FindQuery }

        enum SortingSpec { None, Priority, Hierarchy }

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
        static SortingSpec sortingSpec = SortingSpec.None;
        static List<Task> allTasksInViewFlattened;

        static string usage =
@"Usage:
HansoftExport -c<server>:<port>:<database>:<sdk user>:<pwd> -p<project>:[a|s|b|q] -r<report>:<user>|-f<query> -o:<file>

This utility exports the data of a Hansoft report or a Find query to Excel. All active columns in Hansoft will be
exported regardless of what columns that has been defined to be visible in the report. There is no guaranteed column
order but the order will be the same as long as the set of active columns remain unchanged. If any sorting or grouping
is defined in the report this will also be ignored.

If any parameter values contain spaces, then the parameter value in question need to be double quoted. Colons are not
allowed in parameter values.

Options -c, -p, and -o must always be specified and one of the options -r and-f must also be specified. The -s option
is optional.

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
Note: if the query expression contains double quotes, they should be replaced with single quotes when using this
      utility.

-o Specifies the name of the Excel output file 
<file>         : File name


Examples:
Find all high priority bugs in the project MyProject in the database My Database where the server is running on the
local machine on port 50257, save the output to the file Bugs.xslx:

HansoftExport -clocalhost:50257:""My Database"":sdk:sdk -pMyProject:q -f""Bugpriority='High priority'"" -oBugs.xlsx

Export all items from the report PBL (defined by Manager Jim) in the product backlog of the project MyProject in the
database My Database found on the server running on the local machine at port 50257:

HansoftExport -clocalhost:50257:""My Database"":sdk:sdk -pMyProject:b -rPBL:""Manager Jim"" -oPBL.xlsx

";

        static void Main(string[] args)
        {
            ILogger logger = new ConsoleLogger();
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
                            logger.Exception(e);
                        }
                        finally
                        {
                            SessionManager.Instance.CloseSession();
                        }
                    }
                    else
                    {
                        logger.Warning("Could not open connection to Hansoft");
                    }
                }
                else
                    logger.Information(usage);
            }
            catch (Exception e)
            {
                logger.Exception(e);
            }
        }

        static void DoExport()
        {            
            Project  project = HPMUtilities.FindProject(projectName);
            ProjectView projectView = project.Schedule;
            if (viewType == EHPMReportViewType.AgileBacklog)
                projectView = project.ProductBacklog;
            else if (viewType == EHPMReportViewType.AllBugsInProject)
                projectView = project.BugTracker;
            if (projectView != null)
            {
                string findString;
                if (searchSpec == SearchSpec.Report)
                {
                    User reportUser = HPMUtilities.GetUsers().Find(u => u.Name == reportUserName);
                    if (reportUser == null)
                        throw new ArgumentException("Could not find the user " + reportUserName);
                    HPMReport report = FindReport(projectView.UniqueID, reportUser.UniqueID, reportName);
                    if (report == null)
                        throw new ArgumentException("Could not find the report " + reportName + " for user " + reportUserName + " in project " + projectName);
                    findString = SessionManager.Session.UtilConvertReportToFindString(report, projectView.UniqueID, viewType);
                }
                else
                {
                    findString = findQuery.Replace('\'', '"');
                }

                List<Task> tasks = projectView.Find(findString);

                ExcelWriter excelWriter = new ExcelWriter();
                if (tasks.Count > 0)
                {
                    allTasksInViewFlattened = new List<Task>(projectView.DeepChildren.Cast<Task>());
                    tasks.Sort(CompareItemsByHierarchyIndex);

                    EHPMProjectGetDefaultActivatedNonHidableColumnsFlag flag;
                    if (viewType == EHPMReportViewType.ScheduleMainProject)
                        flag = EHPMProjectGetDefaultActivatedNonHidableColumnsFlag.ScheduledMode;
                    else if (viewType == EHPMReportViewType.AgileMainProject)
                        flag = EHPMProjectGetDefaultActivatedNonHidableColumnsFlag.AgileMode;
                    else
                        flag = EHPMProjectGetDefaultActivatedNonHidableColumnsFlag.None;


                    EHPMProjectDefaultColumn[] nonHidableColumns = SessionManager.Session.ProjectGetDefaultActivatedNonHidableColumns(projectView.UniqueID, flag).m_Columns;
                    if (viewType == EHPMReportViewType.ScheduleMainProject || viewType == EHPMReportViewType.AgileMainProject || viewType == EHPMReportViewType.AgileBacklog)
                    {
                        Array.Resize(ref nonHidableColumns, nonHidableColumns.Length + 1);
                        nonHidableColumns[nonHidableColumns.Length - 1] = EHPMProjectDefaultColumn.ItemStatus;
                    }
                    EHPMProjectDefaultColumn[] activeBuiltinColumns = SessionManager.Session.ProjectGetDefaultActivatedColumns(projectView.UniqueID).m_Columns;
                    HPMProjectCustomColumnsColumn[] activeCustomColumns = SessionManager.Session.ProjectCustomColumnsGet(projectView.UniqueID).m_ShowingColumns;
                      
                    ExcelWriter.Row row = excelWriter.AddRow();

                    foreach (EHPMProjectDefaultColumn builtinCol in nonHidableColumns)
                        row.AddCell(GetColumnName(builtinCol));
                    foreach (EHPMProjectDefaultColumn builtinCol in activeBuiltinColumns)
                        row.AddCell(GetColumnName(builtinCol));
                    foreach (HPMProjectCustomColumnsColumn customColumn in activeCustomColumns)
                        row.AddCell(GetColumnName(projectView.UniqueID, customColumn));

                    HPMColumnTextOptions options = new HPMColumnTextOptions();
                    options.m_bForDisplay = true;
                    foreach (Task item in tasks)
                    {

                        row = excelWriter.AddRow();
                        foreach (EHPMProjectDefaultColumn builtinCol in nonHidableColumns)
                        {
                            HPMColumn column = new HPMColumn();
                            column.m_ColumnType = EHPMColumnType.DefaultColumn;
                            column.m_ColumnID = (uint)builtinCol;
                            row.AddCell(SessionManager.Session.TaskRefGetColumnText(item.UniqueID, column, options));
                        }
                        foreach (EHPMProjectDefaultColumn builtinCol in activeBuiltinColumns)
                        {
                            HPMColumn column = new HPMColumn();
                            column.m_ColumnType = EHPMColumnType.DefaultColumn;
                            column.m_ColumnID = (uint)builtinCol;
                            row.AddCell(SessionManager.Session.TaskRefGetColumnText(item.UniqueID, column, options));
                        }
                        foreach (HPMProjectCustomColumnsColumn customColumn in activeCustomColumns)
                            row.AddCell(item.GetCustomColumnValue(customColumn).ToString());
                    }
                }
                excelWriter.SaveAsOfficeOpenXml(outputFileName);
            }
            else
                throw new ArgumentException("Could not find the project " + projectName);
        }

        public static int CompareItemsByHierarchyIndex(Task t1, Task t2)
        {
            return allTasksInViewFlattened.FindIndex(ii => ii.UniqueID.m_ID == t1.UniqueID.m_ID) - allTasksInViewFlattened.FindIndex(ii => ii.UniqueID.m_ID == t2.UniqueID.m_ID);
        }

        public static HPMReport FindReport(HPMUniqueID projId, HPMUniqueID reportUserId, string reportName)
        {
            HPMReports reports = SessionManager.Session.ProjectGetReports(projId, reportUserId);
            foreach (HPMReport report in reports.m_Reports)
            {
                if (report.m_Name == reportName)
                    return report;
            }
            return null;
        }

        public static string GetColumnName(HPMUniqueID projId, HPMProjectCustomColumnsColumn column)
        {
            return column.m_Name;
        }

        public static string GetColumnName(EHPMProjectDefaultColumn columnId)
        {
            HPMUntranslatedString columnName = SessionManager.Session.UtilGetColumnName(columnId);
            return SessionManager.Session.LocalizationTranslateString(SessionManager.Session.LocalizationGetDefaultLanguage(), columnName);
        }

        static bool ParseArguments(string[] args)
        {
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
