using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using PowerPointLabs;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPExtraEventHelper
{
    internal class PPMouse
    {
        private static int hook;
        private static int _regionHook;

        private static bool isInit = false;

        private static PowerPoint.Selection selectedRange;

        private static IntPtr slideViewWindowHandle;

        private static Rectangle slideViewWindowRectangle;

        private static Native.HookProc hookProcedure;
        private static Native.HookProc _regionHookProcedure;

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

        public delegate void ClickEventDelegate(Point mousePosition);

        //Handler
        public static event DoubleClickEventDelegate DoubleClick;
        public static event ClickEventDelegate Click; 

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

        private static int ClickHookProcedureCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode >= 0)
            {
                if (wParam.ToInt32() == (int)Native.Message.WM_LBUTTONDOWN ||
                    wParam.ToInt32() == (int)Native.Message.WM_LBUTTONUP)
                {
                    var mousePoint = Native.GetPoint(lParam);
                    var mousePos = new Point(mousePoint.x, mousePoint.y);

                    if (Click != null)
                    {
                        Click(mousePos);
                    }
                }
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

        public static void StartRegionClickHook()
        {
            _regionHookProcedure = ClickHookProcedureCallback;
            _regionHook = Native.SetWindowsHookEx((int) Native.HookType.WH_MOUSE, _regionHookProcedure, 0,
                                                  Native.GetWindowThreadProcessId(
                                                      new IntPtr(Globals.ThisAddIn.Application.HWND), 0));
        }

        public static bool StopHook()
        {
            return Native.UnhookWindowsHookEx(hook);
        }

        public static bool StopRegionHook()
        {
            return Native.UnhookWindowsHookEx(_regionHook);
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
}