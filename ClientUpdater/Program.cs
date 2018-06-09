using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;

namespace ClientUpdater
{
    static class Program
    {
        /// <summary>
        /// 해당 응용 프로그램의 주 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main(params string[] parameters)
        {

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            string clientUpdateName = string.Empty;
            string mainPath = string.Empty;
            string clientFileName = string.Empty;
            mainPath = @"C:\HansCreative\VC\SSES_Program";
            mainPath = mainPath.Replace('+', ' ');
            Application.Run(new ClientUpdater(clientUpdateName, mainPath));
        }
    }
}
