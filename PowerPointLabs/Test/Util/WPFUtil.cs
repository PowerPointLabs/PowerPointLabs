using System;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows;
using System.Windows.Interop;

namespace Test.Util
{
    public class WPFUtil
    {
        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        [DllImport("user32.dll")]
        private static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

        [DllImport("user32.dll")]
        private static extern IntPtr GetActiveWindow();

        public static Window GetFrontWindow()
        {
            
            Window w = Application.Current.Windows.OfType<Window>().SingleOrDefault(x => x.IsActive);
            return w;

            /*
            IntPtr handle = GetForegroundWindow();//GetActiveWindow();
            HwndSource hwndSource;
            if (handle == IntPtr.Zero || (hwndSource = HwndSource.FromHwnd(handle)) == null) { return null; }
            Window window = hwndSource.RootVisual as Window;
            return window;
            */
        }

        // using MessageBoxUtil as an example, things to do:
        // 1. Write an async wrapper, ... showDialog
        // 2. look at their wait function for the task
    }
}
