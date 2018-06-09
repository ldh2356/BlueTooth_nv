using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
//using UANGEL.Oam.Core.IO;
using System.Xml.Linq;
using System.IO;
//using UANGEL.Oam.Core.Database;
using System.Windows.Forms;
using System.Data;
//using UANGEL.Oam.Core;
using System.Diagnostics;
//using DevExpress.XtraEditors;
//using MySql.Data.MySqlClient;

//using MySql.Data.MySqlClient;

using MySql.Data.MySqlClient;


namespace SSES_Program
{
    public static class UpdateChecker
    {
        public static MainForm main = null; 
        public static bool NeedUpdate(MainForm pmain)
        {
            try
            {
                main = pmain;
                string connString = string.Empty;
                string currentVersion = string.Empty;

                string startPath = Application.StartupPath;
                string updateConfigPath = Path.Combine(@"C:\HansCreative\VC\SSES_Program", "UpdateInfo.xml");

                if (File.Exists(updateConfigPath) == false)
                    return false;

                XElement root = XElement.Load(updateConfigPath);           
                
                string ipAddress = root.Element("Database").Element("IP").Value;
                string port = root.Element("Database").Element("Port").Value;
                string userId = root.Element("Database").Element("User").Value;
                string password = root.Element("Database").Element("Password").Value;
                string databaseName = root.Element("Database").Element("DatabaseName").Value;
                currentVersion = root.Element("CurrentVersion").Value;

                connString = string.Format("server={0};Port={1};User Id={2}; Password={3}; Database={4}; pooling=true;Charset=euckr;",
                    ipAddress, port, userId, password, databaseName);

                //conn = new MySqlConnection(connString);

                MySqlConnection conn = new MySqlConnection(connString);
                //MySqlDatabaseConnection conn = new MySqlDatabaseConnection(connString);

          
                conn.Open();

                DataSet dataset = MySqlHelper.ExecuteDataset(conn, "SELECT UPDATE_TIME FROM TBL_PKG_STORAGE");

                //DataSet dataset = conn.SqlSelectExecuteUsingDataSet("SELECT UPDATE_TIME FROM CF_pkg_storage WHERE PKG_NAME = 'R092'");

                if (dataset != null)
                {
                    if (dataset.Tables[0].Rows.Count > 0)
                    {
                        DateTime updateTime = dataset.Tables[0].Rows[0].Field<DateTime>("UPDATE_TIME");

                        string timeString = TimeConverter.DateTimeToString(updateTime);

                        if (currentVersion.Length == 0)
                        {
                            //MessageBox.Show("버젼 정보가 없는 경우 무조건 받는다.");
                            return true;
                        }

                        if (currentVersion.CompareTo(timeString) < 0)
                        {
                            //MessageBox.Show("현재 최신 버젼 날짜 :" + currentVersion + "서버 최신 버젼 날짜 : " + timeString);
                            return true;
                        }
                    }
                    else
                        return false; // 현재 업데이트 바이너리가 데이터베이스에 존재하지 않음.
                }
                else
                {
                    return false;
                }
            }
            catch (Exception exception)
            {
                return false;
            }

            return false;
        }

        public static void RunClientUpdater()
        {
            string startPath = Application.StartupPath;
            string clientUpdaterPath = Path.Combine(startPath, "ClientUpdater.exe");
            
            string currentPath = Application.StartupPath;

            currentPath = currentPath.Replace(' ', '+');    //설치된 경로에 space가 있으면 각각 parameters로 인식 하기때문에 replace함.
            string parameters =  " " + currentPath;
            
            try
            {
                ProcessStartInfo startInfo = new ProcessStartInfo(clientUpdaterPath);
                //startInfo.Arguments = currentPath;
                //startInfo.Arguments = siteName;
                startInfo.Arguments = parameters;
                startInfo.UseShellExecute = true;
                startInfo.WorkingDirectory = Path.Combine(currentPath, "ClientUpdater");

                Process.Start(startInfo);

                if (main != null)
                {
                    main.Close();
                    Application.Exit();
                }
                //Application.Exit();              

            }
            catch (Exception e)
            {
                //XtraMessageBox.Show(e.ToString());
            }
        }
    }
}
