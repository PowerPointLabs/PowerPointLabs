using System;
using System.Runtime.InteropServices;
using System.Text;

namespace PPExtraEventHelper
{
    internal class Native
    {
        [DllImport("Gdi32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern int GetPixel(
            System.IntPtr hdc,    // handle to DC
            int nXPos,  // x-coordinate of pixel
            int nYPos   // y-coordinate of pixel
        );

        [DllImport("User32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern IntPtr GetDC(IntPtr wnd);

        [DllImport("User32.dll", CallingConvention = CallingConvention.StdCall)]
        public static extern void ReleaseDC(IntPtr dc);

        //Minimum supported client: Vista
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool AddClipboardFormatListener(IntPtr hwnd);

        //Minimum supported client: Vista
        [DllImport("user32.dll", SetLastError = true)]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool RemoveClipboardFormatListener(IntPtr hwnd);

        //Minimum supported client: Windows 2000
        [DllImport("user32.dll")]
        internal static extern IntPtr SetClipboardViewer(IntPtr hwnd);

        //Minimum supported client: Windows 2000
        [DllImport("user32.dll")]
        internal static extern IntPtr ChangeClipboardChain(IntPtr hwnd, IntPtr hWndNext);

        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Auto)]
        internal static extern int GetWindowThreadProcessId(IntPtr hwnd, int ID);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("user32.dll")]
        [return: MarshalAs(UnmanagedType.Bool)]
        internal static extern bool GetWindowRect(HandleRef hWnd, out RECT lpRect);

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        internal static extern int SendMessage(IntPtr hWnd, uint msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", SetLastError = true)]
        internal static extern bool PostMessage(IntPtr hWnd, uint Msg, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Auto)]
        internal static extern int SetWindowsHookEx(int idHook, HookProc lpfn, int hInstance, int threadId);

        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Auto)]
        internal static extern bool UnhookWindowsHookEx(int idHook);

        [DllImport("user32.dll", CallingConvention = CallingConvention.StdCall, CharSet = CharSet.Auto)]
        internal static extern int CallNextHookEx(int idHook, int nCode, IntPtr wParam, IntPtr lParam);

        [DllImport("user32.dll")]
        internal static extern IntPtr SetWinEventHook(uint eventMin, uint eventMax, IntPtr
           hmodWinEventProc, WinEventDelegate lpfnWinEventProc, uint idProcess,
           uint idThread, uint dwFlags);

        [DllImport("user32.dll")]
        internal static extern bool UnhookWinEvent(IntPtr hWinEventHook);

        [DllImport("winmm.dll")]
        internal static extern int mciSendString(string mciCommand,
                                                StringBuilder mciRetInfo,
                                                int infoLen,
                                                IntPtr callBack);

        internal delegate void WinEventDelegate(IntPtr hWinEventHook, uint eventType,
        IntPtr hwnd, int idObject, int idChild, uint dwEventThread, uint dwmsEventTime);

        internal delegate int HookProc(int nCode, IntPtr wParam, IntPtr lParam);

        [StructLayout(LayoutKind.Sequential)]
        internal struct RECT
        {
            internal int Left;        // x position of upper-left corner
            internal int Top;         // y position of upper-left corner
            internal int Right;       // x position of lower-right corner
            internal int Bottom;      // y position of lower-right corner
        }

        [StructLayout(LayoutKind.Sequential)]
        internal class MouseHookStruct
        {
            internal Point pt;
            internal int hwnd;
            internal int wHitTestCode;
            internal int dwExtraInfo;
        }

        [StructLayout(LayoutKind.Sequential)]
        internal class Point
        {
            internal int x;
            internal int y;
        }

        internal enum HookType
        {
            WH_MOUSE = 0x7,
        }

        internal enum Message
        {
            WM_COMMAND = 0x111,
            WM_LBUTTONDBLCLK = 0x0203,
            WM_DRAWCLIPBOARD = 0x308,
            WM_CHANGECBCHAIN = 0x30D,
            WM_CLIPBOARDUPDATE = 0x031D,
        }

        internal enum Event
        {
            EVENT_SYSTEM_MENUEND = 0x5,
            EVENT_OBJECT_CREATE = 0x8000,
        }
    }
}
