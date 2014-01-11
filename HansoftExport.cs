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
        static DateTime firstDate;
        static DateTime lastDate;
        static bool hOptionFound = false;
        static List<Task> allTasksInViewFlattened;

        static string usage =
@"Usage:
HansoftExport -c<server>:<port>:<database>:<sdk user>:<pwd> -p<project>:(a|s|b|q) -r<report>:<user>|-f<query> [-h<first>[:<last>]] -o:<file>

This utility exports the data of a Hansoft report or a Find query to Excel. 

Unless the -h option is specified, all
active columns in Hansoft will be exported regardless of what columns that has been defined to be visible in the
report. There is no guaranteed column order but the order will be the same as long as the set of active columns
remain unchanged. If any sorting or grouping is defined in the report this will also be ignored.

if the -h option is specified then aggregated historical data vill be exported. Gibven the items returned by the
specified report or find query, the total number of items in each respective state (Not done, In progress,
Completed and so on) will be output for each day between the first/last dates given as parameters to the option.

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

-h Output a history of item states per day
<first>        : The first date to include
<last>         : The last date to include, defaults to the current day.
Note: Dates should be formatted according to your current locale, e.g. {%1} for the current day.

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
                    if (!hOptionFound)
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
                    else
                    {

                        List<StatusAggregate> statusChangeHistory = BuildStatusChangeHistory(tasks, firstDate, lastDate);
                        ExcelWriter.Row row = excelWriter.AddRow();
                        row.AddCell("Date");
                        row.AddCell("Blocked");
                        row.AddCell("Completed");
                        row.AddCell("To be deleted");
                        row.AddCell("In progress");
                        row.AddCell("Not done");
                        foreach (StatusAggregate statusAggregate in statusChangeHistory)
                        {
                            row = excelWriter.AddRow();
                            row.AddCell(statusAggregate.Day.ToShortDateString());
                            row.AddCell(statusAggregate.NumberOfBlockedItems.ToString());
                            row.AddCell(statusAggregate.NumberOfCompletedItems.ToString());
                            row.AddCell(statusAggregate.NumberOfDeletedItems.ToString());
                            row.AddCell(statusAggregate.NumberOfInProgressItems.ToString());
                            row.AddCell(statusAggregate.NumberOfNotDoneItems.ToString());
                        }
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

        class StatusChange
        {
            internal Task Task;
            internal DateTime DateTime;
            internal EHPMTaskStatus ToStatus;
        }

        class StatusAggregate
        {
            internal DateTime Day;
            internal int NumberOfBlockedItems;
            internal int NumberOfCompletedItems;
            internal int NumberOfDeletedItems;
            internal int NumberOfInProgressItems;
            internal int NumberOfNotDoneItems;
        }

        private static List<StatusAggregate> BuildStatusChangeHistory(List<Task> tasks, DateTime firstDay, DateTime lastDay)
        {
            List<StatusAggregate> statusChangeHistory = new List<StatusAggregate>();
            List<StatusAggregate> statusChangeHistoryPruned = new List<StatusAggregate>();
            Dictionary<Task, EHPMTaskStatus> currentTaskStatuses = new Dictionary<Task, EHPMTaskStatus>();
            List<StatusChange> statusChanges = new List<StatusChange>();

            HPMDataHistoryGetHistoryParameters pars = new HPMDataHistoryGetHistoryParameters();
            pars.m_FieldID = EHPMStatisticsField.NoStatistics;
            pars.m_FieldData = 0;
            pars.m_DataIdent0 = EHPMStatisticsScope.NoStatisticsScope;
            pars.m_DataIdent1 = 0;

            // Loop through all once to intiate fetching of the histories from the server
            foreach (Task task in tasks)
            {
                HPMDataHistory history = SessionManager.Session.DataHistoryGetHistory(pars);
                pars.m_DataID = task.UniqueTaskID;
            }

            // Now we start to loop over them again and hope to have data available
            foreach (Task task in tasks)
            {
                int maxAttempts = 100;
                int nAttempts = 0;
                pars.m_DataID = task.UniqueTaskID;
                HPMDataHistory history = SessionManager.Session.DataHistoryGetHistory(pars);
                while (history == null && nAttempts < maxAttempts)
                {
                    System.Threading.Thread.Sleep(100);
                    history = SessionManager.Session.DataHistoryGetHistory(pars);
                    nAttempts += 1;
                }
                for (uint i=0; i<history.m_HistoryEntries.Length; i+=1)
                {
                    HPMDataHistoryEntry entry = history.m_HistoryEntries[i];
                    // Check if it is the status field
                    if (entry.m_FieldID == 15 && (entry.m_EntryType == EHPMDataHistoryEntryType.FieldChanged || entry.m_EntryType == EHPMDataHistoryEntryType.FieldCreated))
                    {
                        if (entry.m_bHasDataRecorded)
                        {
                            StatusChange statusChange = new StatusChange();
                            statusChange.DateTime = HPMUtilities.FromHPMDateTime(entry.m_Time);
                            statusChange.Task = task;
                            HPMVariantData data = SessionManager.Session.DataHistoryGetEntryData(history, i);
                            int value = SessionManager.Session.VariantDecode_HPMInt32(data);
                            EHPMTaskStatus status = (EHPMTaskStatus)value;
                            statusChange.ToStatus = status;
                            statusChanges.Add(statusChange);
                        }
                    }
                }
                currentTaskStatuses[task] = EHPMTaskStatus.NoStatus;
            }

            // Sort the statuschanges from the oldest to the newest
            statusChanges.Sort(new StatusChangeComparer());

            DateTime firstDayWithStatusChange = statusChanges.Count>0 ? statusChanges[0].DateTime.Date : DateTime.MaxValue;
            DateTime lastDayWithStatusChange = statusChanges.Count>0 ? statusChanges[statusChanges.Count - 1].DateTime.Date : DateTime.MinValue;

            DateTime currentDay = DateTime.Now.Date;
            TimeSpan loggedSpan = lastDayWithStatusChange - firstDayWithStatusChange;
            int nLoggedDays = loggedSpan.Days + 1;
            int changeIndex = 0;
            for (DateTime iDate = firstDayWithStatusChange; iDate <= lastDayWithStatusChange; iDate = iDate.AddDays(1))
            {
                while (changeIndex < statusChanges.Count && statusChanges[changeIndex].DateTime.Date == iDate)
                {
                    currentTaskStatuses[statusChanges[changeIndex].Task] = statusChanges[changeIndex].ToStatus;
                    changeIndex +=1;
                }
                StatusAggregate statusAggregate = new StatusAggregate();
                statusAggregate.Day = iDate;
                statusAggregate.NumberOfBlockedItems = currentTaskStatuses.Values.Count(t => t == EHPMTaskStatus.Blocked);
                statusAggregate.NumberOfCompletedItems = currentTaskStatuses.Values.Count(t => t == EHPMTaskStatus.Completed);
                statusAggregate.NumberOfDeletedItems = currentTaskStatuses.Values.Count(t => t == EHPMTaskStatus.Deleted);
                statusAggregate.NumberOfInProgressItems = currentTaskStatuses.Values.Count(t => t == EHPMTaskStatus.InProgress);
                statusAggregate.NumberOfNotDoneItems = currentTaskStatuses.Values.Count(t => t == EHPMTaskStatus.NotDone);
                statusChangeHistory.Add(statusAggregate);
            }
            TimeSpan returnedSpan = lastDay - firstDay;
            int nReturnedDays = returnedSpan.Days+1;

            for (DateTime iDate = firstDay; iDate < firstDayWithStatusChange; iDate = iDate.AddDays(1))
            {
                StatusAggregate statusAggregate = new StatusAggregate();
                statusAggregate.Day = iDate;
                statusAggregate.NumberOfBlockedItems = 0;
                statusAggregate.NumberOfCompletedItems = 0;
                statusAggregate.NumberOfDeletedItems = 0;
                statusAggregate.NumberOfInProgressItems = 0;
                statusAggregate.NumberOfNotDoneItems = 0;
                statusChangeHistoryPruned.Add(statusAggregate);
            }
            foreach (StatusAggregate statusAggregate in statusChangeHistory)
            {
                if (statusAggregate.Day > lastDay)
                    break;
                if (statusAggregate.Day >= firstDay)
                    statusChangeHistoryPruned.Add(statusAggregate);
            }
            if (lastDay > lastDayWithStatusChange)
            {
                StatusAggregate prototype = statusChangeHistory.Last();
                for (DateTime iDate = lastDayWithStatusChange.AddDays(1); iDate <= lastDay; iDate = iDate.AddDays(1))
                {
                    StatusAggregate clone = new StatusAggregate();
                    clone.Day = iDate;
                    clone.NumberOfBlockedItems = prototype.NumberOfBlockedItems;
                    clone.NumberOfCompletedItems = prototype.NumberOfCompletedItems;
                    clone.NumberOfDeletedItems = prototype.NumberOfDeletedItems;
                    clone.NumberOfInProgressItems = prototype.NumberOfInProgressItems;
                    clone.NumberOfNotDoneItems = prototype.NumberOfNotDoneItems;
                    statusChangeHistoryPruned.Add(clone);
                }
            }
            return statusChangeHistoryPruned;
        }


        private class StatusChangeComparer : IComparer<StatusChange>
        {
            public int Compare(StatusChange x, StatusChange y)
            {
                if (x.DateTime < y.DateTime)
                    return -1;
                else if (x.DateTime > y.DateTime)
                    return 1;
                else
                    return 0;
            }
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
                        if (hOptionFound)
                            throw new ArgumentException("The -h option can only be specified once");
                        if (pars.Length < 1 || pars.Length > 2)
                            throw new ArgumentException("The -h option was not specified correctly");
                        if (!DateTime.TryParse(pars[0], out firstDate))
                            throw new ArgumentException("The <first> parameter of the -h option was not specified correctly");
                        if (pars.Length == 2)
                        {
                            if (!DateTime.TryParse(pars[1], out lastDate))
                                throw new ArgumentException("The <last> parameter of the -h option was not specified correctly");
                        }
                        else
                            lastDate = DateTime.Now.Date;
                        hOptionFound = true;
                        break;
                    default:
                        throw new ArgumentException("An unsupported option was specifed: " + option);
                }
            }
            return (cOptionFound && pOptionFound && oOptionFound && (rOptionFound || fOptionFound));
        }
    }
}
