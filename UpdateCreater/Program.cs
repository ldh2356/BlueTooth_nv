using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using System.Threading;
using DevExpress.XtraEditors;

namespace UpdateCreater
{
    static class Program
    {
        /// <summary>
        /// 해당 응용 프로그램의 주 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // 영문 리소스를 사용하기 위해서는 아래 사용해야 함
            Thread.CurrentThread.CurrentCulture = new System.Globalization.CultureInfo("en-US"); // ko-KR, neutral
            Thread.CurrentThread.CurrentUICulture = new System.Globalization.CultureInfo("en-US"); // ko-KR, neutral

            //try
            //{
                Application.Run(new UpdateCreater());
            /*}
            catch (Exception ea)
            {
                XtraMessageBox.Show(string.Format("Error: \n   {0}", ea.Message), "Error");
            }*/
        }
    }
}
