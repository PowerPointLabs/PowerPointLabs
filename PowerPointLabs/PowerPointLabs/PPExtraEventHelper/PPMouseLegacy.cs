using System;
using System.Diagnostics;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPExtraEventHelper
{
    [Obsolete("DO NOT use this class! Instead, use PPMouse.")]
    internal class PPMouseLegacy
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
                    var handle = Process.GetCurrentProcess().MainWindowHandle;
                    FindSlideViewWindowHandle(handle);
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

        //for Office 2010, its window structure is like MDIClient --> mdiClass --> paneClassDC (SlideView)
        //but for Office 2013, it's like MDIClient --> mdiClass (SlideView)
        //this structure can be found using SPY++ provided by visual studio
        private static void FindSlideViewWindowHandle(IntPtr handle)
        {
            IntPtr mdiClient = Native.FindWindowEx(handle, IntPtr.Zero, "MDIClient", "");
            if (mdiClient != IntPtr.Zero)
            {
                IntPtr mdiClass = Native.FindWindowEx(mdiClient, IntPtr.Zero, "mdiClass", "");
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