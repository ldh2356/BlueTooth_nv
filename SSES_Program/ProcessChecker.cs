using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Forms;

namespace SSES_Program
{
    static class ProcessChecker
    {
        /// <summary>
        /// 찾아야 할 캡션
        /// </summary>
        static string _requiredString;

        internal static class NativeMethods
        {
            // 현재 실행중인 윈도우의 상태를 보여줌
            [DllImport("user32.dll")]
            public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);

            // 선택한 윈도우를 뒤에 숨어있었으면 앞으로, 최소화상태였으면 원래상태로 되돌려놓으며 활성화시킨다.
            [DllImport("user32.dll")]
            public static extern bool SetForegroundWindow(IntPtr hWnd);

            // EnumWindows의 실행 결과를 받아줄 콜백함수, EnumWindows는 이 함수 결과가 false가 될 때까지 계속 윈도우를 검색하게 됨
            [DllImport("user32.dll")]
            public static extern bool EnumWindows(EnumWindowsProcDel lpEnumFunc, Int32 lParam);

            // HWND 값을 이용하여 프로세스 ID를 알려줌
            [DllImport("user32.dll")]
            public static extern int GetWindowThreadProcessId(IntPtr hWnd, ref Int32 lpdwProcessId);

            // 윈도우의 캡션을 가져옴
            [DllImport("user32.dll")]
            public static extern int GetWindowText(IntPtr hWnd, StringBuilder lpString, Int32 nMaxCount);

            //윈도우의 상태를 normal로 하게 하는 상수
            public const int SW_SHOWNORMAL = 1;
        }

        public delegate bool EnumWindowsProcDel(IntPtr hWnd, Int32 lParam);

        // 모든 실행중인 윈도우를 검색하며 찾고자 하는 캡션의 윈도우를 발견하면 활성화시킨다.
        static private bool EnumWindowsProc(IntPtr hWnd, Int32 lParam)
        {
            int processId = 0;
            NativeMethods.GetWindowThreadProcessId(hWnd, ref processId);

            StringBuilder caption = new StringBuilder(1024);
            NativeMethods.GetWindowText(hWnd, caption, 1024); //방금 검색한 윈도우의 캡션을 가져온다.

            //찾을 윈도우명과 가져온 캡션이 일치한다면,
            if (processId == lParam && (caption.ToString().IndexOf(_requiredString, StringComparison.OrdinalIgnoreCase) != -1))
            {
                //윈도우를 normal 상태로 바꾸고 제일 앞으로 가져온다.
                NativeMethods.ShowWindowAsync(hWnd, NativeMethods.SW_SHOWNORMAL);
                NativeMethods.SetForegroundWindow(hWnd);
            }
            return true; //왜 계속 true만 반환해야 할까???
        }


        static public bool IsOnlyProcess(string forceTitle)
        {
            _requiredString = forceTitle;

            //먼저 실행파일의 이름으로 이름이 같은 프로세스를 검색해본다.
            foreach (Process proc in Process.GetProcessesByName(Application.ProductName))
            {
                if (proc.Id != Process.GetCurrentProcess().Id)
                {
                    NativeMethods.EnumWindows(new EnumWindowsProcDel(EnumWindowsProc), proc.Id);
                    return false;
                }
            }
            return true;
        }
    }
}
