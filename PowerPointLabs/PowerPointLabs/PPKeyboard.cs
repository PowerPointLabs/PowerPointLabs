using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using PowerPointLabs;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Text;

namespace PPExtraEventHelper
{
    internal class PPKeyboard
    {
        // An action that does nothing.
        private static readonly Action EmptyAction = () => { };

        private static Dictionary<int, bool> Keypressed;
        private static Dictionary<int, Action> KeyDownActions;
        private static Dictionary<int, Action> KeyUpActions;

        private static Native.HookProc hookProcedure;

        private static int hookId;
        private static bool Initialised = false;

        public static void Init(PowerPoint.Application application)
        {
            if (!Initialised)
            {
                Initialised = true;
                InitialiseDictionaries();

                application.WindowSelectionChange += (selection) =>
                {
                    if (!IsHooked())
                    {
                        IntPtr PPHandle = Process.GetCurrentProcess().MainWindowHandle;
                        CreateHook(PPHandle);
                    }
                };
            }
        }

        private static void InitialiseDictionaries()
        {
            Keypressed = new Dictionary<int, bool>();
            KeyDownActions = new Dictionary<int, Action>();
            KeyUpActions = new Dictionary<int, Action>();

            foreach (var key in Enum.GetValues(typeof(Native.VirtualKey)))
            {
                int keyIndex = (int)key;
                Keypressed.Add(keyIndex, false);
                KeyDownActions.Add(keyIndex, EmptyAction);
                KeyUpActions.Add(keyIndex, EmptyAction);
            }
        }

        public static void CreateHook(IntPtr handle)
        {
            var slideViewWindowHandle = FindSlideViewWindowHandle(handle);
            hookProcedure = HookProcedureCallback;
            hookId = Native.SetWindowsHookEx((int)Native.HookType.WH_KEYBOARD, hookProcedure, 0,
                Native.GetWindowThreadProcessId(slideViewWindowHandle, 0));
        }

        public static bool StopHook()
        {
            return Native.UnhookWindowsHookEx(hookId);
        }
        
        public static void AddKeydownAction(Native.VirtualKey key, Action action)
        {
            KeyDownActions[(int) key] += action;
        }

        public static void AddKeyupAction(Native.VirtualKey key, Action action)
        {
            KeyUpActions[(int) key] += action;
        }

        //for Office 2010, its window structure is like MDIClient --> mdiClass --> paneClassDC (SlideView)
        //but for Office 2013, it's like MDIClient --> mdiClass (SlideView)
        //this structure can be found using SPY++ provided by visual studio
        private static IntPtr FindSlideViewWindowHandle(IntPtr handle)
        {
            IntPtr slideViewWindowHandle = IntPtr.Zero;
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
                    //FindSlideViewWindowRectangle();
                }
            }
            return slideViewWindowHandle;
        }

        private static int HookProcedureCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            if (nCode == 0)
            {
                int keyIndex = wParam.ToInt32();
                if (IsKeydownCommand(lParam))
                {
                    if (!Keypressed[keyIndex])
                    {
                        KeyDownActions[keyIndex]();
                        Keypressed[keyIndex] = true;
                    }
                }
                else
                {
                    if (Keypressed[keyIndex])
                    {
                        KeyUpActions[keyIndex]();
                        Keypressed[keyIndex] = false;
                    }
                }
            }
            return Native.CallNextHookEx(0, nCode, wParam, lParam);
        }

        /// <summary>
        /// Returns true when lParam refers to a KeyDown event, an false when it is a KeyUp event.
        /// </summary>
        private static bool IsKeydownCommand(IntPtr lParam)
        {
            // It seems that the first bit of the IntPtr lParam decides whether it is a keyDown or keyUp.
            // 0 => keyDown, 1 => keyUp.
            return lParam.ToInt32() >= 0;
        }

        private static bool IsHooked()
        {
            return hookId != 0;
        }

    }
}