using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Interop;
using System.Windows.Threading;
using HWND = System.IntPtr;

namespace Test.Util.Windows
{

    /// <summary>Contains functionality to get window handles.</summary>
    public static class WindowUtil
    {
        public static uint GetProcessId(int window)
        {
            return GetProcessId(new HWND(window));
        }

        public static uint GetProcessId(HWND window)
        {
            GetWindowThreadProcessId(window, out uint processId);
            return processId;
        }

        public static IDictionary<HWND, string> GetOpenWindows(HWND context)
        {
            return GetOpenWindows(GetProcessId(context));
        }

        /// <summary>Returns a dictionary that contains the handle and title of all the open windows.</summary>
        /// <returns>A dictionary that contains the handle and title of all the open windows.</returns>
        public static IDictionary<HWND, string> GetOpenWindows(uint processId)
        {
            HWND shellWindow = GetShellWindow();
            Dictionary<HWND, string> windows = new Dictionary<HWND, string>();

            EnumWindows(delegate (HWND hWnd, int lParam)
            {
                if (hWnd == shellWindow)
                {
                    return true;
                }
                if (!IsWindowVisible(hWnd))
                {
                    return true;
                }
                GetWindowThreadProcessId(hWnd, out uint result);
                if (result != processId)
                {
                    return true;
                }

                windows[hWnd] = GetWindowTitle(hWnd);
                return true;
            }, 0);

            return windows;
        }

        public static string GetWindowTitle(int hWnd)
        {
            return GetWindowTitle(new HWND(hWnd));
        }

        public static string GetWindowTitle(HWND hWnd)
        {
            int length;
            if (hWnd == HWND.Zero || (length = GetWindowTextLength(hWnd)) == 0)
            {
                return "";
            }

            StringBuilder builder = new StringBuilder(length);
            GetWindowText(hWnd, builder, length + 1);
            return builder.ToString();
        }

        private delegate bool EnumWindowsProc(HWND hWnd, int lParam);

        [DllImport("USER32.DLL", SetLastError = true)]
        private static extern uint GetWindowThreadProcessId(HWND hWnd, out uint lpdwProcessId);

        [DllImport("USER32.DLL")]
        private static extern bool EnumWindows(EnumWindowsProc enumFunc, int lParam);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowText(HWND hWnd, StringBuilder lpString, int nMaxCount);

        [DllImport("USER32.DLL")]
        private static extern int GetWindowTextLength(HWND hWnd);

        [DllImport("USER32.DLL")]
        private static extern bool IsWindowVisible(HWND hWnd);

        [DllImport("USER32.DLL")]
        private static extern HWND GetShellWindow();
    }
}
