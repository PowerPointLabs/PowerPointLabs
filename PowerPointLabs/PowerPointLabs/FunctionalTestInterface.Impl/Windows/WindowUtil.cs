using System.Collections.Generic;
using System.Runtime.InteropServices;
using System.Text;
using HWND = System.IntPtr;

namespace PowerPointLabs.FunctionalTestInterface.Windows
{
    /// <summary>
    /// An utility class for retrieving windows and handles;
    /// </summary>
    public static class WindowUtil
    {
        private const uint WM_CLOSE = 0x0010;

        /// <summary>
        /// Closes a window using its HWND.
        /// </summary>
        /// <param name="hwnd">Window handle</param>
        public static void CloseWindow(HWND hwnd)
        {
            SendMessage(hwnd, WM_CLOSE, HWND.Zero, HWND.Zero);
        }

        public static uint GetProcessId(int window)
        {
            return GetProcessId(new HWND(window));
        }

        /// <summary>
        /// Gets the process id of a window using its HWND.
        /// </summary>
        /// <param name="window">Window handle</param>
        /// <returns>Process id</returns>
        public static uint GetProcessId(HWND window)
        {
            GetWindowThreadProcessId(window, out uint processId);
            return processId;
        }

        /// <summary>
        /// Retrieves a list of open windows with context of a process.
        /// </summary>
        /// <param name="context">HWND in process</param>
        /// <returns>Dictionary of window handles and names</returns>
        public static IDictionary<HWND, string> GetOpenWindows(HWND context)
        {
            return GetOpenWindows(GetProcessId(context));
        }

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

        /// <summary>
        /// Retrieves the title of a window using its HWND.
        /// </summary>
        /// <param name="hWnd">Window handle</param>
        /// <returns>Window title</returns>
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

        [DllImport("user32.dll", CharSet = CharSet.Auto)]
        private static extern HWND SendMessage(HWND hWnd, uint msg, HWND wParam, HWND lParam);
    }
}
