using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Xml.Linq;
using MySql.Data.MySqlClient;
using ICSharpCode.SharpZipLib.Zip;
using ICSharpCode.SharpZipLib.Core;
using System.Diagnostics;
using DevExpress.XtraEditors;

namespace ClientUpdater
{
    public partial class ClientUpdater : Form
    {
        MySqlConnection conn;

        BackgroundWorker bgWorker = new BackgroundWorker();
        string extractPath = string.Empty;
        string downloadPath = string.Empty;
        string clientFileName = string.Empty;
        string updateInfoPath = string.Empty;
        string arguments = string.Empty;

        int bufferSize = 1024;
        byte[] buffer;

        bool updateResult = false;

        string description = string.Empty;
        DateTime updateTime = DateTime.Now;
        string configName = string.Empty;

        public ClientUpdater(string clientUpdateName, string mainPath)
        {

            this.extractPath = mainPath;
            this.configName = clientUpdateName;
            this.downloadPath = Path.Combine(Application.StartupPath, "Update");
            

            InitializeComponent();
        }

        private void ClientUpdater_Load(object sender, EventArgs e)
        {

            buffer = new byte[bufferSize];

            string startPath = @"C:\HansCreative\nv\SSES_Program";

            updateInfoPath = Path.Combine(startPath, "UpdateInfo.xml");

            XElement root = XElement.Load(updateInfoPath);

            /*
            <?xml version="1.0" encoding="utf-8" ?>
            <UpdateInfo>
            <Database>
            <IP>192.168.1.122</IP>
            <Port>3306</Port>
            <DatabaseName>sktscp</DatabaseName>
            <User>omp</User>
            <Password>omp.123</Password>
            </Database>
            <CurrentVersion></CurrentVersion>
            </UpdateInfohttp://m.sports.naver.com/worldbaseball/gamecenter/mlb/index.nhn?tab=&gameId=20150618BOAT0>
            */

            string ipAddress = root.Element("Database").Element("IP").Value;
            string port = root.Element("Database").Element("Port").Value;
            string userId = root.Element("Database").Element("User").Value;
            string password = root.Element("Database").Element("Password").Value;
            string databaseName = root.Element("Database").Element("DatabaseName").Value;

            string connString = string.Format("server={0};Port={1};User Id={2}; Password={3}; Database={4}; pooling=true;Charset=euckr;",
                ipAddress, port, userId, password, databaseName);

            conn = new MySqlConnection(connString);

            try
            {
                conn.Open();
            }
            catch (Exception exception)
            {
                XtraMessageBox.Show(string.Format("There was a critical error. The updator is terminated:\n{0}", exception.Message), "Error");
                Application.Exit();
            }

            DataSet dataset = MySqlHelper.ExecuteDataset(conn, "SELECT DESCRIPTION, UPDATE_TIME FROM TBL_PKG_STORAGE");

            if (dataset == null)
                return;

            if (dataset.Tables[0].Rows.Count == 0)
                return;

            description = dataset.Tables[0].Rows[0].Field<string>("DESCRIPTION");
            updateTime = dataset.Tables[0].Rows[0].Field<DateTime>("UPDATE_TIME");

            //MessageBox.Show("ClientUpdater_Load2");

            bgWorker.DoWork += new DoWorkEventHandler(bgWorker_DoWork);
            bgWorker.ProgressChanged += new ProgressChangedEventHandler(bgWorker_ProgressChanged);
            bgWorker.RunWorkerCompleted += new RunWorkerCompletedEventHandler(bgWorker_RunWorkerCompleted);
            bgWorker.WorkerReportsProgress = true;
        }

        void bgWorker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            if (updateResult)
            {
                XElement root = XElement.Load(updateInfoPath);
                IEnumerable<XElement> updateConfigs = from el in root.Elements("Database") select el;

                root.Element("CurrentVersion").Value = DateTimeToString(updateTime);

                root.Save(updateInfoPath);
                this.downLoadUnpackingWizardPage.AllowNext = true;
            }
            
        }

        void bgWorker_ProgressChanged(object sender, ProgressChangedEventArgs e)
        {
            try
            {
                this.progressBarControl.EditValue = e.ProgressPercentage;
                if (e.UserState != null)
                {
                    this.updateStatusRichTextBox.AppendText(e.UserState.ToString());
                    this.updateStatusRichTextBox.AppendText("\n");
                    this.updateStatusRichTextBox.ScrollToCaret();
                }
            }
            catch
            {

            }
        }

        void bgWorker_DoWork(object sender, DoWorkEventArgs e)
        {
            try
            {
                long readSize = 0;
                long offset = 0;

                int progressPercent = 0;

                bgWorker.ReportProgress(progressPercent, "File Download Begin...");


                string updateFileName = string.Empty;
                updateFileName = "update_tempfile.zip";

                string filePath = Path.Combine(Application.StartupPath, "update");
                //string filePath = Path.Combine(@"C:\HansCreative\VC\SSES_Program", "update");

                if (Directory.Exists(filePath) == false)
                    Directory.CreateDirectory(filePath);

                filePath = Path.Combine(filePath, updateFileName);

                if (File.Exists(filePath) == true)
                {
                    File.Delete(filePath);
                }

                FileStream fs = new FileStream(filePath, FileMode.OpenOrCreate);
                BinaryWriter binaryWriter = new BinaryWriter(fs);

                MySqlCommand cmd = new MySqlCommand("Select LENGTH(BINARY_DATA) FROM TBL_PKG_STORAGE", conn);
                MySqlDataReader reader = cmd.ExecuteReader();

                int binarySize;

                if (reader.Read() == false)
                    return;
                else
                {
                    binarySize = reader.GetInt32(0);
                }
                reader.Close();

                cmd = new MySqlCommand("Select BINARY_DATA FROM TBL_PKG_STORAGE", conn);
                reader = cmd.ExecuteReader();

             

                if (reader.Read() == false)
                    return;
                else
                {
                    readSize = reader.GetBytes(0, 0, buffer, 0, bufferSize);
                    while (readSize == bufferSize)
                    {
                        binaryWriter.Write(buffer);
                        readSize = reader.GetBytes(0, offset, buffer, 0, bufferSize);
                        offset += readSize;

                        progressPercent = (int)(offset * 40 / binarySize);
                        //bgWorker.ReportProgress(progressPercent, string.Format("{0}/{1} 다운 로드 진행 중...", offset, binarySize));
                        bgWorker.ReportProgress(progressPercent);
                    }
                    binaryWriter.Write(buffer, 0, (int)readSize);

                }

                reader.Close();
                binaryWriter.Close();
                fs.Close();

                bgWorker.ReportProgress(progressPercent, "File Download End...");

                ZipFile zipFile = new ZipFile(filePath);

                int itemCount = 0;
                int totalCount = (int)zipFile.Count;
                int percent = progressPercent;


                bgWorker.ReportProgress(progressPercent, "Uncompressing...");
                bgWorker.ReportProgress(progressPercent, string.Format("Uncompress path : {0}", downloadPath));

                try
                {
                    foreach (ZipEntry entry in zipFile)
                    {
                        if (entry.IsFile)
                        {
                            bgWorker.ReportProgress(percent, entry.Name);
                            ExtractFile(zipFile.GetInputStream(entry), entry, extractPath);
                            itemCount++;
                            percent = (int)(itemCount * 60 / totalCount + 40);
                        }

                    }
                }
                catch (Exception exception)
                {
                    updateResult = false;
                    bgWorker.ReportProgress(percent, exception.Message);

                    return;
                }
                bgWorker.ReportProgress(100, string.Format("Unocmpress is completed..."));

                updateResult = true;
            }
            catch(Exception ea)
            {
                MessageBox.Show(ea.Message);
            }

        }


        private void wizardControl1_NextClick(object sender, DevExpress.XtraWizard.WizardCommandButtonClickEventArgs e)
        {
            if (e.Page == this.welcomeWizardPage)
            {
                ShowUpdateDescription();
            }
            else if (e.Page == this.updateDescrptionWizardPage)
            {
                this.downLoadUnpackingWizardPage.AllowNext = false;
                this.downLoadUnpackingWizardPage.AllowBack = false;
                try
                {
                    //MessageBox.Show("RunWorkerAsync 이전");
                    bgWorker.RunWorkerAsync();
                    //MessageBox.Show("RunWorkerAsync 이후");
                }
                catch(Exception ea)
                {
                    MessageBox.Show(ea.Message);
                }
            }

        }

        private void ShowUpdateDescription()
        {
            DataSet dataset = MySqlHelper.ExecuteDataset(conn, "SELECT DESCRIPTION FROM TBL_PKG_STORAGE");

            try
            {
                string description = dataset.Tables[0].Rows[0].Field<string>("DESCRIPTION");
                this.updateInformationRichTextBox.AppendText(description);
            }
            catch (Exception exception)
            {
                CloseUpdater(false, exception.Message);
            }
        }

        private void wizardControl1_CancelClick(object sender, CancelEventArgs e)
        {
            this.Close();
        }

        private void CloseUpdater(bool updated, string msg)
        {
            Close();
        }

        bool ExtractFile(Stream inputStream, ZipEntry theEntry, string targetDir)
        {
            // try and sort out the correct place to save this entry
            string entryFileName;

            if (Path.IsPathRooted(theEntry.Name))
            {
                string workName = Path.GetPathRoot(theEntry.Name);
                workName = theEntry.Name.Substring(workName.Length);
                entryFileName = Path.Combine(Path.GetDirectoryName(workName), Path.GetFileName(theEntry.Name));
            }
            else
            {
                entryFileName = theEntry.Name;
            }

            string targetName = Path.Combine(targetDir, entryFileName);
            string fullPath = Path.GetDirectoryName(Path.GetFullPath(targetName));


            // Could be an option or parameter to allow failure or try creation
            if (Directory.Exists(fullPath) == false)
            {
                try
                {
                    Directory.CreateDirectory(fullPath);
                }
                catch
                {
                    return false;
                }
            }

            if (entryFileName.Length > 0)
            {
                FileStream outputStream = File.Create(targetName);
                StreamUtils.Copy(inputStream, outputStream, buffer);
                outputStream.Close();

            }
            return true;
        }

        private void wizardControl1_FinishClick(object sender, CancelEventArgs e)
        {
            RunClient();
            Close();
        }

        void RunClient()
        {
            Directory.SetCurrentDirectory(downloadPath);

            //string path = string.Format("{0}\\{1}", downloadPath, clientFileName);
            string path = string.Format("{0}\\{1}", @"C:\HansCreative\nv\SSES_Program", "SSES_Program.exe");

            if (File.Exists(path) == true)
            {
                try
                {
                    //XtraMessageBox.Show(string.Format("RunClient {0}", path));
                    ProcessStartInfo startInfo = new ProcessStartInfo(path);
                    startInfo.Arguments = this.arguments;
                    startInfo.WorkingDirectory = extractPath;
                    Process.Start(startInfo);

                }
                catch (Exception e)
                {
                    XtraMessageBox.Show(e.ToString(), "Exception");
                }
            }
        }

        public static string DateTimeToString(DateTime dateTime)
        {
            string time = string.Format("{0,4:0000}-{1,2:00}-{2,2:00} {3,2:00}:{4,2:00}:{5,2:00}"
                        , dateTime.Year, dateTime.Month, dateTime.Day, dateTime.Hour, dateTime.Minute, dateTime.Second);
            return time;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
            /*string startPath = Application.StartupPath;
            string path = Path.Combine(startPath, "SSES_Program.exe");
            ProcessStartInfo startInfo = new ProcessStartInfo(path);
            startInfo.Arguments = this.arguments;
            startInfo.WorkingDirectory = extractPath;
            Process.Start(startInfo);*/
        }

        private void welcomeWizardPage_Click(object sender, EventArgs e)
        {

        }
    }
}
