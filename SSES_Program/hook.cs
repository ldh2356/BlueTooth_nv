using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading;
namespace Hook
{
    internal class BlockingQueue<T>
    {
        Queue<T> que = new Queue<T>();
        Semaphore sem = new Semaphore(0, Int32.MaxValue);
        public void Enqueue(T item)
        {
            lock (que)
            {
                que.Enqueue(item);
            }
            sem.Release();
        }
        public T Dequeue()
        {
            sem.WaitOne();
            lock (que)
            {
                return que.Dequeue();
            }
        }
    }
    public abstract class GlobalHook<T>
    {
        protected int hookHandle = 0;
        private BlockingQueue<T> MessageQueue = new BlockingQueue<T>();
        private Thread MessageThread = null;
        private Object lockObject = new Object();
        private Win32.HookProc HookCallBack = null;
        private bool IsValid = true;
        protected abstract int HookId { get; }

        //-------------------------------------
        // 생성자
        //-------------------------------------
        public GlobalHook()
        {
            this.HookCallBack = new Win32.HookProc(this.CallBack);
            System.Windows.Forms.Application.ApplicationExit += new EventHandler(Application_ApplicationExit);
            MessageThread = new Thread(MessageThreadStart) { Name = "GlobalHook", IsBackground = true };
            MessageThread.Start();
        }

        //-------------------------------------
        // Callback Procedure
        //-------------------------------------
        protected virtual int CallBack(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                T msg = this.ParseLowMessage(wParam, lParam);
                this.MessageQueue.Enqueue(msg);
            }
            return Win32.CallNextHookEx(hookHandle, nCode, wParam, lParam);
        }

        //-------------------------------------
        // On Input Detected
        //-------------------------------------
        protected abstract void OnMessageHooked(T input);
        //-------------------------------------
        // ParseLowMessage
        //-------------------------------------
        protected abstract T ParseLowMessage(IntPtr wParam, IntPtr lParam);

        //-------------------------------------
        // Start Hooking
        //-------------------------------------
        public bool StartHook()
        {
            lock (lockObject)
            {
                if (hookHandle != 0) { return false; }
                hookHandle = Win32.SetWindowsHookEx(this.HookId, HookCallBack, GlobalHook<T>.GetCurrentModuleHandle(), 0);
                return (hookHandle != 0);
            }
        }

        //-------------------------------------
        // Get Current Module Handle
        //-------------------------------------
        private static IntPtr GetCurrentModuleHandle()
        {
            using (Process curProcess = Process.GetCurrentProcess())
            {
                using (ProcessModule curModule = curProcess.MainModule)
                {
                    return Win32.GetModuleHandle(curModule.ModuleName);
                }
            }
        }

        //-------------------------------------
        // Stop Hooking
        //-------------------------------------
        public void StopHook()
        {
            lock (lockObject)
            {
                if (hookHandle == 0) { return; }
                Win32.UnhookWindowsHookEx(hookHandle);
                hookHandle = 0;
            }
        }

        //-------------------------------------
        // MessageThreadStart
        //-------------------------------------
        private void MessageThreadStart()
        {
            while (this.IsValid)
            {
                T msg = MessageQueue.Dequeue();
                OnMessageHooked(msg);
            }
        }

        //-------------------------------------
        // Release Resource
        //-------------------------------------
        public void Dispose()
        {
            this.IsValid = false;
            StopHook();
            System.Windows.Forms.Application.ApplicationExit -= new EventHandler(Application_ApplicationExit);
        }

        //-------------------------------------
        // Release
        //-------------------------------------
        void Application_ApplicationExit(object sender, EventArgs e)
        {
            this.Dispose();
        }
    }
    public class KeyboardHook : GlobalHook<KeyboardHookEventArgs>
    {
        public event EventHandler<KeyboardHookEventArgs> MessageHooked;
        protected override int HookId
        {
            get { return Win32.WH_KEYBOARD_LL; }
        }
        protected override void OnMessageHooked(KeyboardHookEventArgs input)
        {
            if (MessageHooked != null)
            {
                MessageHooked(null, input);
            }
        }
        protected override KeyboardHookEventArgs ParseLowMessage(IntPtr wParam, IntPtr lParam)
        {
            return new KeyboardHookEventArgs(wParam, lParam);
        }
    }
    public class KeyboardHookEventArgs : EventArgs
    {
        [Serializable]
        public enum Messages { WM_KEYDOWN = 0x0100, WM_KEYUP = 0x0101, WM_SYSKEYDOWN = 0x0104, WM_SYSKEYUP = 0x0105 };
        public Messages Message { get; set; }
        public Win32.KBDLLHOOKSTRUCT Data { get; set; }

        //---------------------------------------
        // 생성자
        //---------------------------------------
        internal KeyboardHookEventArgs(IntPtr wParam, IntPtr lParam)
        {
            Win32.KBDLLHOOKSTRUCT kbHook = (Win32.KBDLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(Win32.KBDLLHOOKSTRUCT));
            this.Message = (Messages)wParam;
            this.Data = kbHook;
        }

        public override string ToString()
        {
            return String.Format("Time: {0}, Message: {1}, vkCode: {2}, ScanCode: {3}", this.Data.time, this.Message, this.Data.vkCode, this.Data.scanCode);
        }
    }
    public class MouseHook : GlobalHook<MouseHookEventArgs>
    {
        public event EventHandler<MouseHookEventArgs> MessageHooked;

        protected override int HookId
        {
            get { return Win32.WH_MOUSE_LL; }
        }

        protected override void OnMessageHooked(MouseHookEventArgs input)
        {
            if (MessageHooked != null)
            {
                MessageHooked(null, input);
            }
        }

        protected override MouseHookEventArgs ParseLowMessage(IntPtr wParam, IntPtr lParam)
        {
            return new MouseHookEventArgs(wParam, lParam);
        }
    }
    public class MouseHookEventArgs : EventArgs
    {
        [Serializable]
        private enum Messages
        {
            WM_LBUTTONDOWN = 0x0201, WM_LBUTTONUP = 0x0202, WM_MOUSEMOVE = 0x0200, WM_MOUSEWHEEL = 0x020A, WM_MOUSEHWHEEL = 0x020E, WM_RBUTTONDOWN = 0x0204, WM_RBUTTONUP = 0x0205,
            WM_MBUTTONDOWN = 0x0207, WM_MBUTTONUP = 0x0208
        };
        private Messages Message { get; set; }
        private Win32.MSLLHOOKSTRUCT Data { get; set; }

        internal MouseHookEventArgs(IntPtr wParam, IntPtr lParam)
        {
            Win32.MSLLHOOKSTRUCT msHook = (Win32.MSLLHOOKSTRUCT)Marshal.PtrToStructure(lParam, typeof(Win32.MSLLHOOKSTRUCT));
            this.Message = (Messages)wParam;
            this.Data = msHook;
        }

        public override string ToString()
        {
            return String.Format("Time: {0}, Message: {1}, Point: ({2},{3})", this.Data.time, this.Message, this.Data.point.x, this.Data.point.y);
        }
    }
    public sealed class Win32
    {
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        internal static extern int SetWindowsHookEx(int idHook, HookProc lpfn, IntPtr hInstance, int threadId);
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        internal static extern bool UnhookWindowsHookEx(int idHook);
        [DllImport("user32.dll", CharSet = CharSet.Auto, CallingConvention = CallingConvention.StdCall)]
        internal static extern int CallNextHookEx(int idHook, int nCode, IntPtr wParam, IntPtr lParam);
        [DllImport("kernel32.dll", CharSet = CharSet.Auto)]
        internal static extern IntPtr GetModuleHandle(string lpModuleName);
        internal delegate int HookProc(int nCode, IntPtr wParam, IntPtr lParam);
        [StructLayout(LayoutKind.Sequential), Serializable()]
        public class KBDLLHOOKSTRUCT
        {
            public int vkCode;
            public int scanCode;
            public int flags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential), Serializable()]
        public class POINT
        {
            public uint x;
            public uint y;
        }
        [StructLayout(LayoutKind.Sequential), Serializable()]
        public class MSLLHOOKSTRUCT
        {
            public POINT point;
            public uint mouseData;
            public int flags;
            public int time;
            public IntPtr dwExtraInfo;
        }

        internal const int WH_KEYBOARD_LL = 13;
        internal const int WH_MOUSE_LL = 14;
    }
}
