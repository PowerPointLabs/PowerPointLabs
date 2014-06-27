using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Text;

namespace PPExtraEventHelper
{
    internal class PPMouse
    {
        private static int hook;
        private static bool isInit = false;
        private static PowerPoint.Selection selectedRange;
        private static IntPtr slideViewWindowHandle;
        private static Rectangle slideViewWindowRectangle;
        private static Native.HookProc hookProcedure;
        private static double startTimeInMillisecond = CurrentMillisecond();

        public static void Init(PowerPoint.Application application)
        {
            if (!isInit)
            {
                isInit = true;
                application.WindowSelectionChange += (selection) =>
                {
                    selectedRange = selection;
                    if (!IsHookSuccessful())
                    {
                        IntPtr PPHandle = Process.GetCurrentProcess().MainWindowHandle;
                        StartHook(PPHandle);
                    }
                };
            }
        }

        private static bool IsHookSuccessful()
        {
            return hook != 0;
        }

        //Delegate
        public delegate void DoubleClickEventDelegate(PowerPoint.Selection selection);

        //Handler
        public static event DoubleClickEventDelegate DoubleClick;

        private static int HookProcedureCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                if (wParam.ToInt32() == (int)Native.Message.WM_LBUTTONDBLCLK 
                    && !IsReEnteredCallback())
                {
                    if (IsMouseWithinSlideViewWindow()
                        && DoubleClick != null)
                    {
                        DoubleClick(selectedRange);
                    }
                }
                UpdateStartTime();  
            }
            return Native.CallNextHookEx(0, nCode, wParam, lParam);
        }

        private static void UpdateStartTime()
        {
            startTimeInMillisecond = CurrentMillisecond();
        }

        private static bool IsReEnteredCallback()
        {
            double currentTime = CurrentMillisecond();
            return currentTime - startTimeInMillisecond <= 10;
        }

        private static double CurrentMillisecond()
        {
            return (DateTime.UtcNow - new DateTime(1970, 1, 1)).TotalMilliseconds;
        }

        private static bool IsMouseWithinSlideViewWindow()
        {
            float x = Cursor.Position.X;
            float y = Cursor.Position.Y;
            return x > slideViewWindowRectangle.X 
                && x < slideViewWindowRectangle.X + slideViewWindowRectangle.Width 
                && y > slideViewWindowRectangle.Y 
                && y < slideViewWindowRectangle.Y + slideViewWindowRectangle.Height;
        }

        public static void StartHook(IntPtr handle)
        {
            FindSlideViewWindowHandle(handle);
            hookProcedure = HookProcedureCallback;
            hook = Native.SetWindowsHookEx((int)Native.HookType.WH_MOUSE, hookProcedure, 0, 
                Native.GetWindowThreadProcessId(slideViewWindowHandle, 0));
        }

        public static bool StopHook()
        {
            return Native.UnhookWindowsHookEx(hook);
        }

        //for Office 2010, its window structure is like MDIClient --> mdiClass --> paneClassDC (SlideView)
        //but for Office 2013, it's like MDIClient --> mdiClass (SlideView)
        //this structure can be found using SPY++ provided by visual studio
        private static void FindSlideViewWindowHandle(IntPtr handle)
        {
            IntPtr MDIClient = Native.FindWindowEx(handle, IntPtr.Zero, "MDIClient", "");
            if (MDIClient != IntPtr.Zero)
            {
                IntPtr mdiClass = Native.FindWindowEx(MDIClient, IntPtr.Zero, "mdiClass", "");
                if (mdiClass != IntPtr.Zero)
                {
                    slideViewWindowHandle = Native.FindWindowEx(mdiClass, IntPtr.Zero, "paneClassDC", "Slide");
                    if (slideViewWindowHandle == IntPtr.Zero)
                    {
                        slideViewWindowHandle = mdiClass;
                    }
                    FindSlideViewWindowRectangle();
                }
            }
        }

        private static void FindSlideViewWindowRectangle()
        {
            Native.RECT rec;
            Native.GetWindowRect(new HandleRef(new object(), slideViewWindowHandle), out rec);
            slideViewWindowRectangle = new Rectangle();
            slideViewWindowRectangle.X = rec.Left;
            slideViewWindowRectangle.Y = rec.Top;
            slideViewWindowRectangle.Width = rec.Right - rec.Left + 1;
            slideViewWindowRectangle.Height = rec.Bottom - rec.Top + 1;
        }        
    }

    public class SysMouseEventInfo : EventArgs
    {
        public string WindowTitle { get; set; }
    }
    public class LMouseUpListener
    {

        int HC_ACTION = 0;
        Native.HookProc CallBack = null;
        IntPtr _hook = IntPtr.Zero;

        public event EventHandler<SysMouseEventInfo> LButtonUpClicked;
        public LMouseUpListener()
        {
            this.CallBack += new Native.HookProc(MouseEvents);
            using (System.Diagnostics.Process process = System.Diagnostics.Process.GetCurrentProcess())
            using (System.Diagnostics.ProcessModule module = process.MainModule)
            {
                IntPtr hModule = Native.GetModuleHandle(module.ModuleName);
                _hook = Native.SetWindowsHookEx(
                    (int)Native.HookType.WH_MOUSE_LL, 
                    this.CallBack, 
                    hModule, 
                    0);
            }
        }

        int MouseEvents(int code, IntPtr wParam, IntPtr lParam)
        {
            if (code < 0)
                return Native.CallNextHookEx(_hook, code, wParam, lParam);

            if (code == this.HC_ACTION)
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
    }
}