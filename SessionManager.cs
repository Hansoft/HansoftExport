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
    class SessionManager
    {
        private static SessionManager instance = null;

        private string sdkUser;
        private string sdkUserPwd;
        private string server;
        private int port;
        private string database;

        HPMSdkSession hpmSession;
        HPMSdkCallback hpmCallback;

        public static void Initialize(string sdkUser, string sdkUserPwd, string server, int port, string database)
        {
            if (instance != null)
                instance.CloseSession();
            instance = new SessionManager(sdkUser, sdkUserPwd, server, port, database);
        }

        private SessionManager()
        {
        }

        public static SessionManager Instance
        {
            get
            {
                if (instance == null)
                    instance = new SessionManager();

                return instance;
            }
        }

        private SessionManager(string sdkUser, string sdkUserPwd, string server, int port, string database)
        {
            this.sdkUser = sdkUser;
            this.sdkUserPwd = sdkUserPwd;
            this.server = server;
            this.port = port;
            this.database = database;
            hpmCallback = new HPMSdkCallback();
        }

        public void CloseSession()
        {
            if (hpmSession != null)
                HPMSdkSession.SessionDestroy(ref hpmSession);
        }

        public bool Connected
        {
            get
            {
                return hpmSession != null;
            }
        }

        public bool Connect()
        {
            if (hpmSession == null)
            {
                try
                {
                    hpmSession = HPMSdkSession.SessionOpen(server, port, database, sdkUser, sdkUserPwd, hpmCallback, null, true, EHPMSdkDebugMode.Debug, (IntPtr)null, 0, "", "", null);
                    Logger.Information("Connection established");
                }
                catch (Exception e)
                {
                    Logger.Exception("Could not connect to Hansoft Server", e);
                }
            }
            return hpmSession!=null;
        }

        public HPMSdkSession Session
        {
            get
            {
                if (hpmSession == null)
                    Connect();
                return hpmSession;
            }
        }

        class HPMSdkCallback : HPMSdkCallbacks
        {
            public override void On_ProcessError(EHPMError error)
            {
                Logger.Error("On_ProcessError: " + HPMSdkSession.ErrorToStr(error));
            }
        }

    }
}
