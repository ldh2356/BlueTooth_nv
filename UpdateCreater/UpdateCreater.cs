using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using MySql.Data.MySqlClient;
using System.IO;
using System.Xml.Linq;
using DevExpress.XtraEditors;

namespace UpdateCreater
{
    public partial class UpdateCreater : Form
    {
        MySqlConnection conn;

        public UpdateCreater()
        {
            InitializeComponent();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            OpenFileDialog dlg = new OpenFileDialog();

            if (dlg.ShowDialog() == DialogResult.OK)
            {
                this.textEdit1.Text = dlg.FileName;
            }
        }

        private void btnUpload_Click(object sender, EventArgs e)
        {

            int result = MySqlHelper.ExecuteNonQuery(conn, "DELETE  FROM TBL_PKG_STORAGE");

            if (result < 0)
            {
                XtraMessageBox.Show("An error found in uploading the package.", "Error");
                return;
            }

            string filePath = this.textEdit1.Text;

            string description =  string.Format("{0}\n\n\n{1}", this.newVersionRichTextBox.Text, this.historyRichTextBox.Text);


            MySqlDataAdapter da = new MySqlDataAdapter("Select * from TBL_PKG_STORAGE", conn);

            DataSet ds = new DataSet("binaryData");

            MySqlCommandBuilder MyCb = new MySqlCommandBuilder(da);
            da.MissingSchemaAction = MissingSchemaAction.AddWithKey;
            da.Fill(ds, "binaryData");

            DataRow newRow = ds.Tables["binaryData"].NewRow();


            FileStream fs = new FileStream(filePath, FileMode.Open);
            BinaryReader br = new BinaryReader(fs);
            byte[] bytes = new byte[fs.Length];
            br.Read(bytes, 0, (int)fs.Length);

            br.Close();
            fs.Close();

     
            newRow["BINARY_DATA"] = bytes;
            newRow["UPDATE_TIME"] = DateTime.Now;
            newRow["DESCRIPTION"] = description;

            ds.Tables["binaryData"].Rows.Add(newRow);
            da.Update(ds, "binaryData");

            XtraMessageBox.Show("The package has been uploaded successfully.", "Uploading");
            this.Enabled = true;  
        }

        private void UpdateCreater_Load(object sender, EventArgs e)
        {
//            try
//            {
                string startPath = Application.StartupPath;
                string updateInfoPath = Path.Combine(startPath, "UpdateInfo.xml");

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

                DataSet dataset = MySqlHelper.ExecuteDataset(conn, "SELECT DESCRIPTION FROM TBL_PKG_STORAGE");

                if (dataset == null)
                    return;

                if (dataset.Tables[0].Rows.Count == 0)
                    return;

                this.historyRichTextBox.Text = dataset.Tables[0].Rows[0].Field<string>("DESCRIPTION");
//            }
//            catch (Exception ea)
//            {
//               XtraMessageBox.Show(string.Format("Error: \n   {0}", ea.Message), "Error");
//            }

        }

        private void UpdateCreater_FormClosing(object sender, FormClosingEventArgs e)
        {
            conn.Close();
        }
    }
}
