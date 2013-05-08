using HPMSdk;

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

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

namespace Hansoft.SdkUtils
{
    class HPMUtilities
    {
        private HPMUtilities()
        {
        }

        public static HPMUniqueID FindProject(string projectName)
        {
            HPMProjectEnum projects = SessionManager.Instance.Session.ProjectEnum();

            foreach (HPMUniqueID projId in projects.m_Projects)
            {
                HPMProjectProperties props = SessionManager.Instance.Session.ProjectGetProperties(projId);
                if (props.m_Name == projectName)
                    return projId;
            }
            return null;
        }

        public static HPMUniqueID FindUser(string userName)
        {
            HPMResourceEnum users = SessionManager.Instance.Session.ResourceEnum();

            foreach (HPMUniqueID userId in users.m_Resources)
            {
                HPMResourceProperties props = SessionManager.Instance.Session.ResourceGetProperties(userId);
                if (props.m_Name == userName)
                    return userId;
            }
            return null;
        }

        public static string GetUserName(HPMUniqueID id)
        {
            HPMResourceEnum users = SessionManager.Instance.Session.ResourceEnum();
            foreach (HPMUniqueID userId in users.m_Resources)
            {
                if (userId == id)
                {
                    HPMResourceProperties props = SessionManager.Instance.Session.ResourceGetProperties(userId);
                    return props.m_Name;
                }
            }
            return null;
        }

        public static string GetGroupName(HPMUniqueID id)
        {
            HPMResourceGroupEnum groups = SessionManager.Instance.Session.ResourceGroupEnum();
            foreach (HPMUniqueID groupId in groups.m_ResourceGroups)
            {
                if (groupId == id)
                {
                    HPMResourceGroupProperties props = SessionManager.Instance.Session.ResourceGroupGetProperties(groupId);
                    return props.m_Name;
                }
            }
            return null;
        }

        public static HPMReport FindReport(HPMUniqueID projId, HPMUniqueID reportUserId, string reportName)
        {
            HPMReports reports = SessionManager.Instance.Session.ProjectGetReports(projId, reportUserId);
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
            HPMUntranslatedString columnName = SessionManager.Instance.Session.UtilGetColumnName(columnId);
            return SessionManager.Instance.Session.LocalizationTranslateString(SessionManager.Instance.Session.LocalizationGetDefaultLanguage(), columnName);
        }

        internal static HPMString DecodeDroplistValue(HPMString dbValue, HPMProjectCustomColumnsColumnDropListItem[] droplistItem)
        {
            if (dbValue != "")
            {
                int iVal = Int32.Parse(dbValue);
                for (int i = 0; i < droplistItem.Length; i += 1)
                {
                    if (droplistItem[i].m_Id == iVal)
                        return droplistItem[i].m_Name;
                }
                return "Unknown";
            }
            else return "";
        }
    }
}
