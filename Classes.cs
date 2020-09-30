using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.ServiceProcess;
using System.Xml;
using System.Xml.XPath;

namespace iParser
{
    public partial class iParser : ServiceBase
    {
        
    }
    class Config
        {
            string strLoadDirectoryPath;
            string strBackupDirectoryPath;
            string strBadXMLDirectoryPath;
            int intSaveTime;
            string strSQLServerName;
            string strSQLInstanceName;
            string strSQLBaseName;

            public Config()
            {
                
            }

            public string LoadDirectoryPath
            {
                get
                {
                    return strLoadDirectoryPath;
                }
                set
                {
                    strLoadDirectoryPath = value;
                }
            }
            public string BackupDirectoryPath
            {
                get
                {
                    return strBackupDirectoryPath;
                }
                set
                {
                    strBackupDirectoryPath = value;
                }
            }
            public string BadXMLDirectoryPath
            {
                get
                {
                    return strBadXMLDirectoryPath;
                }
                set
                {
                    strBadXMLDirectoryPath = value;
                }
            }
            public int SaveTime
            {
                get
                {
                    return intSaveTime;
                }
                set
                {
                    intSaveTime = value;
                }
            }

            public string SQLServerName
            {
                get
                {
                    return strSQLServerName;
                }
                set
                {
                    strSQLServerName = value;
                }
            }
            public string SQLInstanceName
            {
                get
                {
                    return strSQLInstanceName;
                }
                set
                {
                    strSQLInstanceName = value;
                }
            }
            public string SQLBaseName
            {
                get
                {
                    return strSQLBaseName;
                }
                set
                {
                    strSQLBaseName = value;
                }
            }
        }
    }

    

