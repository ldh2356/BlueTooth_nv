using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Security.Principal;

using System.Net.Sockets;
using InTheHand.Net.Bluetooth;
using InTheHand.Windows.Forms;
using InTheHand.Net.Sockets;
using InTheHand.Net.Bluetooth.AttributeIds;

namespace SSES_Program
{
    static class Program
    {
        /// <summary>
        /// 해당 응용 프로그램의 주 진입점입니다.
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (ProcessChecker.IsOnlyProcess(Application.ProductName))
            {
                try
                {
                    IntPtr ptr = FindWindow(null, "MainForm");
                    if (ptr != IntPtr.Zero)
                    {
                        SendMessage(ptr, 0x0010, 0, 0);
                        System.Environment.Exit(0);
                    }
                    Application.EnableVisualStyles();
                    Application.SetCompatibleTextRenderingDefault(false);
                    Application.Run(new MainForm());
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.ToString() + "Program");
                }
            }
        }


        [DllImport("user32.dll", SetLastError = true)]
        static extern IntPtr FindWindow(string lpClassName, string lpWindowName);
        [DllImport("user32.dll", SetLastError = true)]
        public static extern int SendMessage(IntPtr hwnd, int msg, int wparam, int lparam);
    }
}
