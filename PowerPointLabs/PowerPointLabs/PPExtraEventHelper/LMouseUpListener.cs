using System;
using System.Runtime.InteropServices;
using System.Text;

namespace PPExtraEventHelper
{
    [Obsolete("DO NOT use this class! Instead, use PPMouse.")]
    public class LMouseUpListener
    {

        int hcAction = 0;
        Native.HookProc callBack = null;
        IntPtr _hook = IntPtr.Zero;

        public event EventHandler<SysMouseEventInfo> LButtonUpClicked;
        public LMouseUpListener()
        {
            this.callBack += new Native.HookProc(MouseEvents);
            using (System.Diagnostics.Process process = System.Diagnostics.Process.GetCurrentProcess())
            using (System.Diagnostics.ProcessModule module = process.MainModule)
            {
                IntPtr hModule = Native.GetModuleHandle(module.ModuleName);
                _hook = Native.SetWindowsHookEx(
                    (int)Native.HookType.WH_MOUSE_LL,
                    this.callBack,
                    hModule,
                    0);
            }
        }

        public void Close()
        {
            if (_hook != IntPtr.Zero)
            {
                Native.UnhookWindowsHookEx(_hook);
            }
        }
        public static string GetWindowTextRaw(IntPtr hwnd)
        {
            // Allocate correct string length first
            int length = (int)Native.SendMessage(
                hwnd,
                (int)Native.Message.WM_GETTEXTLENGTH,
                IntPtr.Zero,
                IntPtr.Zero);

            StringBuilder sb = new StringBuilder(length);
            Native.SendMessage(
                hwnd,
                (int)Native.Message.WM_GETTEXT,
                (IntPtr)sb.Capacity,
                sb);

            return sb.ToString();
        }

        int MouseEvents(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code < 0)
            {
                return Native.CallNextHookEx(_hook, code, wParam, lParam);
            }

            if (code == this.hcAction)
            {
                // Left button pressed somewhere
                if (wParam.ToInt32() == (uint)Native.Message.WM_LBUTTONUP)
                {
                    Native.MSLLHOOKSTRUCT ms = new Native.MSLLHOOKSTRUCT();
                    ms = (Native.MSLLHOOKSTRUCT)Marshal.PtrToStructure(
                        lParam,
                        typeof(Native.MSLLHOOKSTRUCT));

                    IntPtr win = Native.WindowFromPoint(ms.pt);

                    string title = GetWindowTextRaw(win);
                    if (LButtonUpClicked != null)
                    {
                        LButtonUpClicked(
                            this,
                            new SysMouseEventInfo { WindowTitle = title });
                    }
                }
            }
            return Native.CallNextHookEx(_hook, code, wParam, lParam);
        }
    }
}
