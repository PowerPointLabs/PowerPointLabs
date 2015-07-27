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
        private static Dictionary<int, bool> Keypressed;
        private static Dictionary<int, List<Func<bool>>> KeyDownActions;
        private static Dictionary<int, List<Func<bool>>> KeyUpActions;

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
            KeyDownActions = new Dictionary<int, List<Func<bool>>>();
            KeyUpActions = new Dictionary<int, List<Func<bool>>>();

            foreach (var key in Enum.GetValues(typeof(Native.VirtualKey)))
            {
                int keyIndex = (int)key;
                Keypressed.Add(keyIndex, false);
                KeyDownActions.Add(keyIndex, new List<Func<bool>>());
                KeyUpActions.Add(keyIndex, new List<Func<bool>>());
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
            AddKeydownAction(key, ReturnFalse(action));
        }

        public static void AddKeydownAction(Native.VirtualKey key, Func<bool> action)
        {
            KeyDownActions[(int)key].Add(action);
        }

        public static void AddKeyupAction(Native.VirtualKey key, Action action)
        {
            AddKeyupAction(key, ReturnFalse(action));
        }

        public static void AddKeyupAction(Native.VirtualKey key, Func<bool> action)
        {
            KeyUpActions[(int)key].Add(action);
        }

        private static Func<bool> ReturnFalse(Action action)
        {
            return () =>
            {
                action();
                return false;
            };
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
            bool blockInput = false;
            if (nCode == 3)
            {
                int keyIndex = wParam.ToInt32();
                if (Keypressed.ContainsKey(keyIndex))
                {
                    if (IsKeydownCommand(lParam))
                    {
                        if (!Keypressed[keyIndex])
                        {
                            foreach (var action in KeyDownActions[keyIndex])
                            {
                                var block = action();
                                if (block) blockInput = true;
                            }
                            Keypressed[keyIndex] = true;
                        }
                    }
                    else
                    {
                        if (Keypressed[keyIndex])
                        {
                            foreach (var action in KeyUpActions[keyIndex])
                            {
                                var block = action();
                                if (block) blockInput = true;
                            }
                            Keypressed[keyIndex] = false;
                        }
                    }
                }
            }
            Debug.WriteLine(blockInput);

            if (blockInput) return 1;
            else return Native.CallNextHookEx(0, nCode, wParam, lParam);
        }

        /// <summary>
        /// Returns true when lParam refers to a KeyDown event, an false when it is a KeyUp event.
        /// </summary>
        private static bool IsKeydownCommand(IntPtr lParam)
        {
            // It seems that the first bit of the IntPtr lParam decides whether it is a keyDown or keyUp.
            // Note: using lParam.ToInt32() here causes an OverflowException on 64-bit machines for some reason.
            return (lParam.ToInt64() & 0x80000000) == 0;
        }

        private static bool IsHooked()
        {
            return hookId != 0;
        }

    }
}