using System;
using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Diagnostics;

namespace Test.Util
{
    using HWND = IntPtr;

    /// <summary>Contains functionality to get all the open windows.</summary>
    public static class OpenWindowGetter
    {
        public static void GetAllChildrenWindowHandles(HWND hParent)
        {
            HWND prevChild = HWND.Zero;
            HWND currChild = HWND.Zero;
            // get number of children windows
            while (true)
            {
                currChild = FindWindowEx(hParent, prevChild, null, null);
                if (currChild == HWND.Zero)
                {
                    break;
                }
                System.Windows.MessageBox.Show(GetWindowText(currChild));
                prevChild = currChild;
            }
        }

        public static void Nuke()
        {
            Process[] processlist = Process.GetProcesses();

            foreach (Process process in processlist)
            {
                if (!String.IsNullOrEmpty(process.MainWindowTitle))
                {
                    System.Windows.MessageBox.Show(string.Format("Process: {0} ID: {1} Window title: {2}", process.ProcessName, process.Id, process.MainWindowTitle));
                }
            }
        }

        private static string GetWindowText(HWND hWnd)
        {
            int length = GetWindowTextLength(hWnd);
            if (length == 0) return "";

            StringBuilder builder = new StringBuilder(length);
            GetWindowText(hWnd, builder, length + 1);
            return builder.ToString();
        }

        /// <summary>Returns a dictionary that contains the handle and title of all the open windows.</summary>
        /// <returns>A dictionary that contains the handle and title of all the open windows.</returns>
        public static IDictionary<HWND, string> GetOpenWindows(HWND parent)
        {
            uint processId;
            GetWindowThreadProcessId(parent, out processId);

            HWND shellWindow = GetShellWindow();
            Dictionary<HWND, string> windows = new Dictionary<HWND, string>();

            EnumWindows(delegate (HWND hWnd, int lParam)
            {
                if (hWnd == shellWindow) return true;
                if (!IsWindowVisible(hWnd)) return true;
                uint result;
                GetWindowThreadProcessId(hWnd, out result);
                if (result != processId) return true;

                int length = GetWindowTextLength(hWnd);
                if (length == 0) return true;

                StringBuilder builder = new StringBuilder(length);
                GetWindowText(hWnd, builder, length + 1);

                windows[hWnd] = builder.ToString();
                return true;

            }, 0);

            return windows;
        }

        private delegate bool EnumWindowsProc(HWND hWnd, int lParam);

        [DllImport("user32.dll", SetLastError = true)]
        static extern uint GetWindowThreadProcessId(HWND hWnd, out uint lpdwProcessId);

        [DllImport("USER32.DLL")]
        private static extern bool EnumWindows(EnumWindowsProc enumFunc, int lParam);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowText(HWND hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowTextLength(HWND hWnd);

        [DllImport("USER32.DLL")]
        private static extern bool IsWindowVisible(HWND hWnd);

        [DllImport("USER32.DLL", EntryPoint = "FindWindowEx", CharSet = CharSet.Auto)]
        static extern HWND FindWindowEx(HWND hwndParent, HWND hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("USER32.DLL")]
        private static extern HWND GetShellWindow();
    }
}
