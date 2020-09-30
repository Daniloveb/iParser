using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
//using System.Diagnostics.
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Xml;
using System.Xml.XPath;
using Microsoft.Win32;
using System.Timers;
using System.Data.SqlClient;
using System.Data.Sql;
using Microsoft.SqlServer;
using Microsoft.SqlServer.Management.Common;
using Microsoft.SqlServer.Server;
using Microsoft.SqlServer.Management.Smo;
using System.IO;
using System.Reflection;

namespace iParser
{
    public partial class iParser : ServiceBase
    {
        static bool bGlobalFailure;
        static bool processing;
        static string connectionString;
        static SqlConnection connection;
        static int iNumber;
        static Guid iGUID;
        static bool iNumberOK;
        static string strPrefix;
        static string strDay;
        static int intLogDay;
        static int intDay;
        string strSPath;
        Config outConfig;
        Server server;
        Database db;
        public iParser()
        {
            InitializeComponent();
        }

        protected override void OnStart(string[] args)
        {
            //Раздел Сервиса
            bGlobalFailure = false;
            //Single Interval = 30000;
            processing = false;
            System.Timers.Timer timer = new System.Timers.Timer(60000);
            timer.Elapsed += new ElapsedEventHandler(timer_Elapsed);
            timer.Start();
            GC.KeepAlive(timer);
            
            //Логирование
            strDay = DateTime.Today.ToString().Remove(10);
            intLogDay = DateTime.Now.DayOfYear;
            strSPath = Path.GetDirectoryName(Assembly.GetEntryAssembly().Location);
            TextWriterTraceListener traceListener = new TextWriterTraceListener(Console.Out);
            Trace.Listeners.Add(traceListener);
            Trace.AutoFlush = true;
            //Проверяем наличие Log Directory
            if (!System.IO.Directory.Exists(strSPath + "\\Logs\\"))
            {
                System.IO.Directory.CreateDirectory(strSPath + "\\Logs\\");
            }
            TextWriterTraceListener traceFileListener = new TextWriterTraceListener(strSPath + "\\Logs\\" + strDay + ".log");
            Trace.Listeners.Add(traceFileListener);
            Trace.WriteLine(DateTime.Now.ToString() + ": Service started successfully!");
            if (!System.Diagnostics.EventLog.SourceExists("iParcer"))
            { System.Diagnostics.EventLog.CreateEventSource("iParcer", "Application"); }
            //Server server;
            try
            {
                //Проверяем существование конфигурационного файла
                if (!System.IO.File.Exists(strSPath + "\\iParserConfig.xml"))
                {
                    Trace.WriteLine(DateTime.Now.ToString() + ": Error! Configurations file not exist!");
                    bGlobalFailure = true;
                    Environment.Exit(255);
                }
                //читаем конфигурационный файл
                outConfig = new Config();
                ReadConfFile(strSPath + "\\iParserConfig.xml", out outConfig);

                //создаем соединение
                connection = new SqlConnection();
                db = new Database();
                try
                {
                    connectionString = "Data Source=" + outConfig.SQLServerName + ";Initial Catalog=" + outConfig.SQLBaseName + ";Integrated Security = true;";
                    connection = new SqlConnection(connectionString);
                    server = new Server(new ServerConnection(connection));
                    //db = server.Databases[outConfig.SQLBaseName];
                }
                catch (Exception e)
                {
                    Trace.WriteLine(DateTime.Now.ToString() + ": DataBase connect Error! " + e.Message);
                    Environment.Exit(255);
                }
                //Проверяем наличие BackupDirectory
                if (!System.IO.Directory.Exists(outConfig.BackupDirectoryPath))
                {
                    System.IO.Directory.CreateDirectory(outConfig.BackupDirectoryPath);
                    Trace.WriteLine("Create Backup Directory!");
                }
                //Проверяем наличие BadXMLDirectory
                if (!System.IO.Directory.Exists(outConfig.BadXMLDirectoryPath))
                {
                    System.IO.Directory.CreateDirectory(outConfig.BadXMLDirectoryPath);
                    Trace.WriteLine("Create Bad XML Directory!");
                }
                //LogClear(outConfig);
                //Parsing(outConfig, db);
            }
            catch (Exception e)
            {
                Trace.WriteLine(DateTime.Now.ToString() + ": Error! " + e.Message);
            }
        }

        void timer_Elapsed(object sender, ElapsedEventArgs e)
        {

            //Сравниваем даты лога и текущей
            intDay = DateTime.Now.DayOfYear;
            if (intDay > intLogDay)
            {
                //создаем новый лог
                strDay = DateTime.Today.ToString().Remove(10);
                TextWriterTraceListener traceFileListener = new TextWriterTraceListener(strSPath + "\\Logs\\" + strDay + ".log");
                Trace.Listeners.Clear();
                Trace.Listeners.Add(traceFileListener);
                intLogDay = intDay;
            }
            Trace.WriteLine(DateTime.Now.ToString() + ": Tick, processing =  " + processing);

            //В случае процесса обработки - не запускаем новый поток
            if (!processing)
            {
                processing = true;

                //Проверяем состояние SQL соединения
                connection = new SqlConnection();
                db = new Database();
                try
                {
                    connectionString = "Data Source=" + outConfig.SQLServerName + ";Initial Catalog=" + outConfig.SQLBaseName + ";Integrated Security = true;";
                    connection = new SqlConnection(connectionString);
                    server = new Server(new ServerConnection(connection));
                    db = server.Databases[outConfig.SQLBaseName];

                    //SqlConnection conn = new SqlConnection("mydatasource");
                    //try
                    //{
                    //    Trace.WriteLine(DateTime.Now.ToString() + " conn.state =  " + connection.State.ToString());
                    //    connection.Open();
                    //    db = server.Databases[outConfig.SQLBaseName];


                    //Trace.WriteLine(DateTime.Now.ToString() + "! " + server.Status.ToString());
                    //Trace.WriteLine(DateTime.Now.ToString() + ": processing value - " + processing.ToString());
                    //string strDay = DateTime.Today.ToString().Remove(10);



                    //Запускаем парсинг файлов
                    try
                    {
                        if (!bGlobalFailure)
                        {
                            LogClear(outConfig);
                            Parsing(outConfig, db);
                            processing = false;
                        }
                    }
                    catch (Exception ex2)
                    {
                        Trace.WriteLine(DateTime.Now.ToString() + ": Error in timerElapsed! " + ex2.Message);
                    }
                }


                catch (Exception ex)
                {
                    Trace.WriteLine(DateTime.Now.ToString() + ": DataBase connect Error! Missing Timer Circle!" + ex.Message);
                    processing = false;
                }
                connection.Close();
                connection.Dispose();
            }
        }

        protected override void OnStop()
        {
            Trace.WriteLine(DateTime.Now.ToString() + ": Error! Service stopped!");
        }

        static void LogClear(Config outConfig)
        {
            DirectoryInfo di = new DirectoryInfo(outConfig.BackupDirectoryPath);
            FileInfo[] fi;
            fi = di.GetFiles();
            foreach (FileInfo currentfile in fi)
            {
                try
                {
                    int m = DateTime.Now.DayOfYear - currentfile.CreationTime.DayOfYear;
                    //Trace.WriteLine("Now.DayOfYear " + DateTime.Now.DayOfYear);
                    //Trace.WriteLine("currentfile.CreationTime.DayOfYear " + currentfile.CreationTime.DayOfYear);
                    //Trace.WriteLine("Имя файла - разница " + currentfile.Name + " - " + m);
                    //Trace.WriteLine(DateTime.Now.ToString() + ":values m + outConfig.SaveTime" + m + " - " + outConfig.SaveTime);
                    if (m > outConfig.SaveTime)
                    {
                        currentfile.Delete();
                        Trace.WriteLine("deleteted ok" + currentfile.Name);
                    }
                }
                catch (Exception e)
                {
                    Trace.WriteLine(DateTime.Now.ToString() + ": Error! Can't delete old Log file" + currentfile.Name + ". " + e.Message);
                }
            }
        }

        static void Parsing(Config outConfig, Database db)
        {
            DirectoryInfo di = new DirectoryInfo(outConfig.LoadDirectoryPath);
            FileInfo[] fi;
            string filename = "";
            strPrefix = "";
            int Nfiles = 0;
            FileInfo file = null;

            fi = di.GetFiles();

            //перебираем файлы
            foreach (FileInfo currentfile in fi)
            {
                try
                {
                    Trace.WriteLine(DateTime.Now.ToString() + ": Parsing Invent File " + currentfile.Name + " ! ");
                    filename = currentfile.Name;
                    file = currentfile;
                    strPrefix = filename.Substring(0, 1);
                    //Проверяем расширение файла, обрабатываем только .xml
                    if (!filename.Substring(6).Equals("xml"))
                    { throw new Exception("Расширение файла не xml!"); }
                    //Парсим инвентарный номер из имени файла
                    int.TryParse(currentfile.Name.Substring(1, 4), out iNumber);
                    iNumberOK = iNumberquery(iNumber, out iGUID);
                    Nfiles++;
                    if (!iNumberOK)// если в БД не найдена запись обрабатываемого инвентарного номера
                    {
                        //Процедура добавляющая запись с таблицу iNumbers
                        AddNumber(iNumber, strPrefix);
                    }
                    //Повторяем запрос в таблицу iNumber
                    iNumberOK = iNumberquery(iNumber, out iGUID);
                    if (!iNumberOK)
                    {
                        Trace.WriteLine(DateTime.Now.ToString() + ": Error! Invent File Name" + filename + " not parsed!");
                    }
                    else
                    {
                        XmlDocument doc = new XmlDocument();
                        doc.Load(currentfile.FullName);
                        XPathNavigator nav = doc.CreateNavigator();
                        XPathNodeIterator it = (XPathNodeIterator)nav.Evaluate("InventoryData");
                        while (it.MoveNext())
                        {
                            if (it.Current is IHasXmlNode)
                            {
                                XmlNode Settingsnode = ((IHasXmlNode)it.Current).GetNode();
                                //Classes
                                foreach (XmlNode XmlNodeClass in Settingsnode.ChildNodes)
                                {
                                    CreateSQLCommand(db, XmlNodeClass);
                                }
                            }
                        }
                        //Перемещаем обработанный файл
                        string strFileNewName;
                        strFileNewName = strDay + "_" + currentfile.Name;
                        if (File.Exists(outConfig.BackupDirectoryPath + strFileNewName)) { File.Delete(outConfig.BackupDirectoryPath + strFileNewName); }
                        currentfile.MoveTo(outConfig.BackupDirectoryPath + strFileNewName);
                    }
                }
                catch (Exception e)
                {
                    Trace.WriteLine(DateTime.Now.ToString() + ": Error! Invent File " + filename + " parsing problem! " + e.Message);
                    if (File.Exists(outConfig.BadXMLDirectoryPath + filename)) { File.Delete(outConfig.BadXMLDirectoryPath + filename); }
                    file.MoveTo(outConfig.BadXMLDirectoryPath + filename);
                }
            }
      
      
        }
        
        // Задача процедуры CreateSQLCommand - за один цикл просмотра файла
        // в случае отсутствия - создать таблицу
        // в случае отсутствия - создать поле
        // создать и запустить INSERT COMMAND
        static void CreateSQLCommand(Database db, XmlNode XmlNodeClass)
        {
            double dval;
            bool bval;
            DateTime dateval;
            bool c_exist;
            string val;
            SqlCommand command;
            string strTableName = XmlNodeClass.Name;
            Trace.WriteLine(DateTime.Now.ToString() + ": Processing " + strTableName + " ! ");
            //Проверяем существование таблицы
            bool bexist = CheckExistSQLTable(db, strTableName);
            Table newTable = new Table(db, strTableName);
            //в случае отсутствия таблицы - добавляем предопределенные поля
            if (!bexist)
            {
                // Add "ID" Column
                Column IDColumn = new Column(newTable, "UID");
                IDColumn.DataType = DataType.UniqueIdentifier;
                IDColumn.Nullable = false;
                IDColumn.RowGuidCol = true;

                // Add "INVNumberIDColumn" Column
                Column INVNumberIDColumn = new Column(newTable, "INVNumberID");
                INVNumberIDColumn.DataType = DataType.UniqueIdentifier;
                INVNumberIDColumn.Nullable = false;

                // Add "Datetime" Column
                Column DateColumn = new Column(newTable, "Date");
                DateColumn.DataType = DataType.DateTime;
                newTable.Columns.Add(IDColumn);
                newTable.Columns.Add(INVNumberIDColumn);
                newTable.Columns.Add(DateColumn);

            }
            string strNameArray = "";
            string strValueArray = "";
            string strINSERT = "";
            //Перебираем экземпляры класса
            foreach (XmlNode XmlNodeInstance in XmlNodeClass.ChildNodes)
            {
                strNameArray = "";
                strValueArray = "";
                strINSERT = "INSERT INTO \"" + strTableName + "\" ([INVNumberID], ";
                //Перебираем аттрибуты экземпляра
                foreach (XmlAttribute XmlAtt in XmlNodeInstance.Attributes)
                {
                    //В случае пустого значения не обрабатываем аттрибут
                    if (XmlAtt.Value != "")
                    {
                        val = XmlAtt.Value;
                        val = val.ToLower();
                        //Обработки предопределенных значений полей
                        if (val == "истина") { val = "true"; }
                        if (val == "ложь") { val = "false"; }
                        Column NewCol = new Column(newTable, XmlAtt.Name);
                        //Обработка поля InstalledOn класса QuickFixEngineering
                        //Меняем местами число и месяц
                        if (XmlAtt.Name.Equals("InstalledOn"))
                        {
                            if (val.Contains("/"))
                            {
                                int sch, sch2;
                                string ttt, uuu, ooo;
                                sch = val.IndexOf("/");
                                sch2 = val.IndexOf("/", sch + 1);
                                ttt = val.Substring(sch + 1, sch2 - sch - 1);
                                uuu = val.Substring(0, sch);
                                ooo = val.Substring(sch2 + 1);
                                val = val.Substring(sch + 1, sch2 - sch - 1) + "-" + val.Substring(0, sch) + "-" + val.Substring(sch2 + 1);
                            }
                            //if (!DateTime.TryParse(val, out dateval))
                            //{
                            //    Debug.WriteLine("not date " + val);
                            //}
                        }
                        //Обработка MAC-адреса - заменяем двоеточия дефисами
                        if (XmlAtt.Name.Equals("MACAddress"))
                        {
                            val = val.Replace(":", "-");
                        }
                        //Определяем Тип данных
                        NewCol.DataType = DataType.VarCharMax;
                        if (Boolean.TryParse(val, out bval))
                        {
                            NewCol.DataType = DataType.Bit;
                        }
                        else if (Double.TryParse(val, out dval))
                        {
                            NewCol.DataType = DataType.Float;
                            //в случае наличия запятых неправильно определяется формат float
                            //Принудительно устанавливаем Varcharmax
                            //Так же в случае больших чисел - за которые ошибочно принимаются серийные номера и т.д.
                            if (val.Contains(",") || val.Length > 20)
                            { NewCol.DataType = DataType.VarCharMax; }
                        }
                        else if (DateTime.TryParse(val, out dateval))// & val.Length == 10)
                        {

                            NewCol.DataType = DataType.Date;
                            //изменяем строку даты под формат SQL c 28.02.2012 на 2012-02-28
                            val = dateval.Year + "-" + dateval.Month + "-" + dateval.Day;
                        }
                        if (bexist) //если таблица существует - проверяем существование поля - смотрим тип данных поля
                        {
                            Table t = db.Tables[strTableName];
                            c_exist = false;
                            if (t.Columns.Contains(XmlAtt.Name))
                            {
                                Column c = t.Columns[XmlAtt.Name];
                                //В случае несовпадения данных SQL поля с типом данных аттрибута - логируем ошибку
                                if (!c.DataType.Name.Equals(NewCol.DataType.Name))
                                {
                                    //В поле varchar можем писать любое значение
                                    //При других несовпадениях обнуляем значение атрибута
                                    if (c.DataType.ToString() != "varchar")
                                    {
                                        Trace.WriteLine("Error!  In column " + c.Name + " with datatype " + c.DataType.ToString() + " in table " + t.Name + ", cant write value " + val + ", with datatype " + NewCol.DataType.ToString() + ". Write Null.");
                                        val = "";
                                    }
                                    else
                                    {
                                        //Trace.WriteLine("Warning!  In column " + c.Name + " with datatype " + c.DataType.ToString() + " in table " + t.Name + ", write value " + val + ", with datatype " + NewCol.DataType.ToString());
                                    }
                                    c_exist = true;
                                }
                                c_exist = true;
                            }
                            //создаем поле
                            if (!c_exist)
                            {
                                Column AddColumn = new Column(db.Tables[strTableName], XmlAtt.Name);
                                AddColumn.DataType = NewCol.DataType;
                                db.Tables[strTableName].Columns.Add(AddColumn);
                                db.Tables[strTableName].Alter();
                            }
                        }
                        else
                        { newTable.Columns.Add(NewCol); }
                        //Убираем "'" в значении. Добавляем ' '
                        if (NewCol.DataType.ToString() == "varchar" || NewCol.DataType.ToString() == "bit" || NewCol.DataType.ToString() == "date")
                        {
                            val = val.Replace("'", "");
                            val = "\'" + val + "\'";
                        }
                        //Заменяем ,  в числовых значениях
                        if (val.Contains(",") & NewCol.DataType.ToString() == "float")
                        {
                            val = val.Replace(",", ".");
                        }
                        bool g = db.QuotedIdentifiersEnabled;
                        //Заполняем INSERT строку
                        if (strNameArray == "")
                        {
                            strNameArray = "[" + XmlAtt.Name + "]";
                            strValueArray = val;
                        }
                        else
                        {
                            strNameArray = strNameArray + ", [" + XmlAtt.Name + "]";
                            strValueArray = strValueArray + ", " + val;
                        }

                        //return strMain + strNameArray + ") VALUES (" + strValueArray + ")";
                        //Trace.WriteLine ("strINSERT " + strINSERT);

                        if (!bexist)
                        {
                            // Создаем индекс
                            Index index = new Index(newTable, "ID" + strTableName);
                            index.IndexKeyType = IndexKeyType.DriPrimaryKey;
                            index.IndexedColumns.Add(new IndexedColumn(index, "UID"));

                            // Добавляем индекс в таблицу
                            newTable.Indexes.Add(index);
                            newTable.Create();
                            newTable.ChangeSchema("dbo");

                            //Определяем SQL Defaults
                            Default def;
                            Default NewID;
                            if (!db.Defaults.Contains("DefDate"))
                            {
                                def = new Default(db, "DefDate");
                                def.TextHeader = "CREATE DEFAULT [DefDate] AS";
                                def.TextBody = "getdate()";
                                def.Create();
                            }
                            if (!db.Defaults.Contains("NUID"))
                            {
                                NewID = new Default(db, "NUID");
                                NewID.TextHeader = "CREATE DEFAULT [NUID] AS";
                                NewID.TextBody = "NewID()";
                                NewID.Create();
                            }
                            def = db.Defaults["DefDate"];
                            def.BindToColumn(strTableName, "Date");

                            NewID = db.Defaults["NUID"];
                            NewID.BindToColumn(strTableName, "UID");
                            bexist = true;
                        }
                    }
                }
                if (!strNameArray.Equals(""))
                {
                    strINSERT = strINSERT + strNameArray + ") VALUES ('" + iGUID.ToString() + "' ," + strValueArray + ")";
                    command = new SqlCommand();
                    command.CommandType = CommandType.Text;
                    command.CommandText = strINSERT;
                    command.Connection = connection;
                    command.ExecuteNonQuery();
                    command.Dispose();
                }
            }
        }
        //Процедура добавляющая запись в таблицу iNumbers
        static void AddNumber(int iNumber, string strPrefix)
        {
            Guid iGuid = Guid.Empty;
            SqlCommand icommand = new SqlCommand();
            icommand.CommandType = CommandType.Text;
            icommand.CommandText = "Select UID from _Types where Prefix ='" + strPrefix + "'";
            //'" + iGuid.ToString() + "'
            //icommand.CommandText = "Select UID from iNumbers where INumber =" + iNumber.ToString();
            icommand.Connection = connection;

            //делаем запрос в базу по префиксу и получаем GUID типа устройства
            try
            {
                iGuid = (Guid)icommand.ExecuteScalar();
            }
            catch (Exception ex)
            {
                Trace.WriteLine(DateTime.Now.ToString() + ": Error! Not founded Prefix " + strPrefix + " in Table Types!" + ex.Message);
            }
            icommand.Dispose();
            //добавляем запись в таблицу iNumbers 
            icommand = new SqlCommand();
            icommand.CommandType = CommandType.Text;
            //icommand.CommandText = "Select UID from Types where INumber =" + Prefix;
            icommand.CommandText = "Insert into _iNumbers (TypeID, iNumber) Values ('" + iGuid.ToString() + "','" + iNumber.ToString() + "')";
            icommand.Connection = connection;
            icommand.ExecuteScalar();
            icommand.Dispose();
        }
        static bool iNumberquery(int iNumber, out Guid iGuid)
        {
            bool r = true;
            iGuid = Guid.Empty;
            SqlCommand icommand = new SqlCommand();
            icommand.CommandType = CommandType.Text;
            icommand.CommandText = "Select UID from _iNumbers where INumber =" + iNumber.ToString();
            icommand.Connection = connection;
            try
            {
                iGuid = (Guid)icommand.ExecuteScalar();
            }
            catch
            {
                r = false;
            }
            icommand.Dispose();
            return r;
        }
        static bool CheckExistSQLTable(Database db, string TableName)
        {
            bool r = false;
            //проверяем на существование таблицы
            foreach (Table t in db.Tables)
            {
                if (t.Name == TableName)
                {
                    r = true;
                }
            }
            return r;
        }
        static void ReadConfFile(string path, out Config outConfig)
        {
            outConfig = new Config();
            XmlDocument doc = new XmlDocument();
            try
            {
                doc.Load(path);
                XPathNavigator nav = doc.CreateNavigator();
                //Считываем Settings
                XPathNodeIterator it = (XPathNodeIterator)nav.Evaluate("ParserConfig/Settings");
                while (it.MoveNext())
                {
                    if (it.Current is IHasXmlNode)
                    {
                        XmlNode Settingsnode = ((IHasXmlNode)it.Current).GetNode();
                        foreach (XmlAttribute XmlAtt in Settingsnode.Attributes)
                        {
                            switch (XmlAtt.Name)
                            {
                                case "LoadDirectoryPath":
                                    outConfig.LoadDirectoryPath = XmlAtt.Value;
                                    break;
                                case "SQLServerName":
                                    outConfig.SQLServerName = XmlAtt.Value;
                                    break;
                                case "SQLInstanceName":
                                    outConfig.SQLInstanceName = XmlAtt.Value;
                                    break;
                                case "SQLBaseName":
                                    outConfig.SQLBaseName = XmlAtt.Value;
                                    break;
                                case "BackupDirectoryPath":
                                    outConfig.BackupDirectoryPath = XmlAtt.Value;
                                    break;
                                case "BadXMLDirectoryPath":
                                    outConfig.BadXMLDirectoryPath = XmlAtt.Value;
                                    break;
                                case "SaveLogDataTime":
                                    outConfig.SaveTime = int.Parse(XmlAtt.Value);
                                    break;

                            }
                        }
                    }
                }
            }
            catch
            {
                Trace.WriteLine(DateTime.Now.ToString() + ": Error in configuration file!");
                //Environment.Exit(255);
                bGlobalFailure = true;
            }
        }
    }
}
