using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Windows.Forms;
using System.Media;
using Microsoft.Win32;


namespace SSES_Program
{
    public class KeyboardHooking
    {
        //FormScreenSaver formScreenSaver;
        /// <summary>
        /// 작업표시줄 숨기기/보이기
        /// </summary>
        /// 상수선언
        public const int SWP_HIDEWINDOW = 128;
        public const int SWP_SHOWWINDOW = 64;
        public static int WINDOWSTATUS;

        //win32 dll 함수 선언**********************************************
        [DllImport("user32.dll")]
        public static extern int FindWindowA(string lpClassName, string lpWindowName);

        [DllImport("user32.dll")]
        public static extern int SetWindowPos(int hwnd, int hWndInsertAfter, int x, int y, int cx, int cy, int wFlags);

        [DllImport("user32")]
        public static extern void LockWorkStation();
        //win32 dll 함수 선언**********************************************

        public static void TaskBarHide()
        {
            int lRet;
            lRet = FindWindowA("Shell_traywnd", "");
            if (lRet > 0)
            {
                lRet = SetWindowPos(lRet, 0, 0, 0, 0, 0, SWP_HIDEWINDOW);
                WINDOWSTATUS = SWP_HIDEWINDOW;
            }
        }

        public static void TaskBarShow()
        {
            int lRet;
            lRet = FindWindowA("Shell_traywnd", "");
            if (lRet > 0)
            {
                lRet = SetWindowPos(lRet, 0, 0, 0, 0, 0, SWP_SHOWWINDOW);
                WINDOWSTATUS = SWP_SHOWWINDOW;
            }
        }

        /// <summary>
        /// 작업관리자 실행시 실행이 안되도록 한다
        /// </summary>
        public static void BlockCtrlAltDel()
        {
            try
            {
                ToggleTaskManagerBlock(false);
            }
            catch (Exception ex)
            {
                MainForm.log.write(ex.StackTrace);
            }
        }

        /// <summary>
        /// 작업관리자 실행시 실행이 안되도록 하는것을 복원한다
        /// </summary>
        public static void UnBlockCtrlAltDel()
        {
            try
            {
                ToggleTaskManagerBlock(true);
            }
            catch (Exception ex)
            {
                MainForm.log.write(ex.StackTrace);
            }
        }



        /// <summary>
        /// 작업관리자 Enable/Disable 을 관리한다
        /// </summary>
        /// <param name="enable"></param>
        public static void ToggleTaskManagerBlock(bool enable)
        {
            try
            {
                RegistryKey objRegistryKey = Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Windows\CurrentVersion\Policies\System");

                // 작업관리자를 복원하고 레지스트리에 DisableTaskMgr 가 등록 안되어있다면 값을 제거
                if (enable && objRegistryKey.GetValue("DisableTaskMgr") != null)
                {
                    objRegistryKey.DeleteValue("DisableTaskMgr");
                }
                // 작업관리자를 Diabled 시킬경우 레지스트리에 값을 추가
                else
                {
                    objRegistryKey.SetValue("DisableTaskMgr", "1");
                }

                objRegistryKey.Close();
            }
            catch (Exception ex)
            {
                MainForm.log.write(ex.StackTrace);
            }
        }

        #region Flag

        const int WM_KEYFIRST = 0x100;
        const int WM_KEYDOWN = 0x100;
        const int WM_KEYUP = 0x101;
        const int WM_CHAR = 0x102;
        const int WM_IME_STARTCOMPOSITION = 0x10D;
        const int WM_IME_COMPOSITION = 0x10F;
        const int WM_IME_ENDCOMPOSITION = 0x10E;
        const int WM_IME_SETCONTEXT = 0x281;
        const int WM_IME_NOTIFY = 0x282;
        const int WM_IME_CONTROL = 0x283;
        const int WM_IME_COMPOSITIONFULL = 0x284;
        const int WM_IME_SELECT = 0x285;
        const int WM_IME_CHAR = 0x286;
        const int WM_IME_KEYDOWN = 0x290;
        const int WM_IME_KEYUP = 0x291;
        const int WM_IME_REPORT = 0x0280;
        const int WM_IME_REQUEST = 0x0288;
        const int WM_SYSKEYDOWN = 0x0104;
        const int WM_SYSKEYUP = 0x0105;

        const int VK_TAB = 0x09;
        const int VK_ESCAPE = 0x1b;
        const int VK_LWIN = 0x5b;
        const int VK_RWIN = 0x5c;
        const int VK_DELETE = 0x2E;

        const int KF_EXTENDED = 0x0100;
        const int KF_ALTDOWN = 0x2000;
        const int LLKHF_EXTENDED = (KF_EXTENDED >> 8);
        const int LLKHF_ALTDOWN = (KF_ALTDOWN >> 8);

        const int WH_KEYBOARD_LL = 13;

        #endregion

        #region DllImport

        [DllImport("user32.dll", EntryPoint = "SetWindowsHookExA", CharSet = CharSet.Ansi)]
        public static extern int SetWindowsHookEx(int idHook, LowLevelKeyboardProcDelegate lpfn, IntPtr hMod, uint dwThreadId);

        [DllImport("user32.dll", EntryPoint = "UnhookWindowsHookEx", CharSet = CharSet.Ansi)]
        public static extern int UnHookWindowsEx(int hHook);

        [DllImport("user32.dll", EntryPoint = "CallNextHookEx", CharSet = CharSet.Ansi)]
        public static extern int CallNextHookEx(int hHook, int nCode, int wParam, ref KBDLLHOOKSTRUCT lParam);

        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
         public static extern IntPtr GetModuleHandle(string lpModuleName);

        #endregion
        
        public delegate int LowLevelKeyboardProcDelegate(int nCode, int wParam, ref KBDLLHOOKSTRUCT lParam);

        public static LowLevelKeyboardProcDelegate hookProc = LowLevelKeyboardProc;

        public static int LowLevelKeyboardProc(int nCode, int wParam, ref KBDLLHOOKSTRUCT lParam)
        {
            bool blnEat = false;
            switch (wParam)
            {
                case WM_KEYDOWN:
                case WM_KEYUP:
                case WM_SYSKEYDOWN:
                case WM_SYSKEYUP:

                    //if (lParam.vkCode == VK_ESCAPE) // ESC 후킹
                    //{
                        //FormScreenSaverCancel.instance.Show();
                    //}

                    //Alt+Tab, Alt+Esc, Ctrl+Esc, Windows Key, ESC
                    if (((lParam.vkCode == VK_TAB) && (lParam.flags == LLKHF_ALTDOWN)) ||
                        ((lParam.vkCode == VK_ESCAPE) && (lParam.flags == LLKHF_ALTDOWN)) ||
                        ((lParam.vkCode == VK_ESCAPE) && (lParam.flags == 0)) ||
                        ((lParam.vkCode == VK_DELETE)) ||
                        ((lParam.vkCode == VK_LWIN) && (lParam.flags == LLKHF_EXTENDED)) ||
                        ((lParam.vkCode == VK_RWIN) && (lParam.flags == LLKHF_EXTENDED)) ||
                        ((true) && (lParam.flags == LLKHF_ALTDOWN)) ||
                        (lParam.vkCode == VK_TAB) ||
                        (lParam.flags == LLKHF_ALTDOWN))
                    {
                        blnEat = true;
                    }
                    //Service.AlertSoundStart();
                    break;
            }

            if (blnEat)
                return 1;
            else return CallNextHookEx(0, nCode, wParam, ref lParam);
        }
        public static void sound()
        {
            //Console.Beep(3000, 10000);
            SystemSounds.Beep.Play();
        }
        public static int SetHook(LowLevelKeyboardProcDelegate proc)
        {
            using (Process curProcess = Process.GetCurrentProcess())
            using (ProcessModule curModule = curProcess.MainModule)
            {
                return SetWindowsHookEx(WH_KEYBOARD_LL, proc, GetModuleHandle(curModule.ModuleName), 0);
            }
        }

        public struct KBDLLHOOKSTRUCT
        {
            public int vkCode;
            int scanCode;
            public int flags;
            int time;
            int dwExtraInfo;
        }
    }
}
