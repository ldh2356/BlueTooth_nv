using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.Threading;//.Tasks;
using System.Runtime.InteropServices;
using System.Media;

namespace SSES_Program
{
    public partial class FormScreenSaver : Form
    {
        int intLLKey;
        public MainForm main;
        public FormScreenSaverCancel formScreenSaverCancel;

        [DllImport("user32.dll")]
        private static extern bool AnimateWindow(IntPtr hWnd, int time, AnimateWindowFlags flags);

        // 플래그 값
        public enum AnimateWindowFlags
        {
            AW_HOR_POSITIVE = 0x00000001,
            AW_HOR_NEGATIVE = 0x00000002,
            AW_VER_POSITIVE = 0x00000004,
            AW_VER_NEGATIVE = 0x00000008,
            AW_CENTER = 0x00000010,
            AW_HIDE = 0x00010000,
            AW_ACTIVATE = 0x00020000,
            AW_SLIDE = 0x00040000,
            AW_BLEND = 0x00080000
        }

        public FormScreenSaver(MainForm main)
        {
            InitializeComponent();
            this.main = main;
            main.SetFormScreenSaver(this);

            FormBorderStyle = System.Windows.Forms.FormBorderStyle.None;
            StartPosition = FormStartPosition.Manual;
            TopMost = true;
            ShowInTaskbar = false;

            pb_screenSaver.BringToFront();

            // 키보드 후킹
            intLLKey = KeyboardHooking.SetHook(KeyboardHooking.hookProc);
            KeyboardHooking.BlockCtrlAltDel();
        }

        // 폼 로드
        private void FormScreenSaver_Load(object sender, EventArgs e)
        {
            // 폼 애니메이션(위에서 아래로)
            AnimateWindow(this.Handle, 500, AnimateWindowFlags.AW_VER_POSITIVE);
        }

        // 폼 액티베이티드
        private void FormScreenSaver_Activated(object sender, EventArgs e)
        {
            KeyboardHooking.TaskBarHide(); // 작업표시줄 숨김
        }

        // 폼 클로즈
        private void FormScreenSaver_FormClosing(object sender, FormClosingEventArgs e)
        {
            // 폼 애니메이션(아래서 위로)
            AnimateWindow(this.Handle, 500,
                AnimateWindowFlags.AW_VER_NEGATIVE | AnimateWindowFlags.AW_HIDE);

            if (KeyboardHooking.WINDOWSTATUS == KeyboardHooking.SWP_HIDEWINDOW)
            {
                KeyboardHooking.TaskBarShow(); // 작업표시줄 드러냄
            }

            // 키보드 후킹 해제
            KeyboardHooking.UnHookWindowsEx(intLLKey);
            KeyboardHooking.UnBlockCtrlAltDel();

            main.rcvRssi = default(int);

            this.Dispose();
        }

        private void FormScreenSaver_KeyDown(object sender, KeyEventArgs e)
        {
        }

        public void SetFormScreenSaverCancel(FormScreenSaverCancel formScreenSaverCancel)
        {
            //this.formScreenSaverCancel = formScreenSaverCancel;
        }

        private void pb_screenSaver_MouseDown(object sender, MouseEventArgs e)
        {
            MainForm.log.write("모니터 1번에 마우스 다운 이벤트");
            formScreenSaverCancel = new FormScreenSaverCancel(this);
            formScreenSaverCancel.TopMost = true;
            formScreenSaverCancel.ShowDialog();
        }

        private void pb_screenSaver_Click(object sender, EventArgs e)
        {
            MainForm.log.write("모니터 1번에 클릭 이벤트");
            formScreenSaverCancel = new FormScreenSaverCancel(this);
            formScreenSaverCancel.TopMost = true;
            formScreenSaverCancel.ShowDialog();
        }
    }
}
