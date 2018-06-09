
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace SSES_Program
{
    public class Log
    {
        private StreamWriter sw = null;

        public void write(String str)
        {
            string FilePath = Application.StartupPath + @"\Log\Log_" + DateTime.Today.ToString("yyyyMMdd") + ".log";
            string DirPath = Application.StartupPath + @"\Log";
            string temp;

            

            DirectoryInfo di = new DirectoryInfo(DirPath);
            FileInfo fi = new FileInfo(FilePath);

            try
            {
                if (di.Exists != true) Directory.CreateDirectory(DirPath);

                if(fi.Exists != true)
                {
                    using (sw = new StreamWriter(FilePath))
                    {
                        temp = string.Format("[{0}] : {1}", GetDateTime(), str );
                        sw.WriteLine(temp);
                    }
                }
                else
                {
                    using (sw = new StreamWriter(FilePath))
                    {
                        temp = string.Format("[{0}] : {1}", GetDateTime(), str);
                        sw.WriteLine(temp);
                        sw.Close();
                    }
                }
            }
            catch(Exception ea)
            {
                sw.Close();
            }
        }

        private string GetDateTime()
        {
            DateTime NowDate = DateTime.Now;
            return NowDate.ToString("yyyy-MM-dd HH:mm:ss") + ":" + NowDate.Millisecond.ToString("000");
        }
    }
}