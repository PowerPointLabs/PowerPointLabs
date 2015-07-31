using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Drawing;
using System.Linq;
using PowerPointLabs;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using System.Text;

namespace PPExtraEventHelper
{
    internal class PPKeyboard
    {
        private static Dictionary<int, KeyStatus> KeyStatuses;
        private static Dictionary<int, List<BindedAction>> KeyDownActions;
        private static Dictionary<int, List<BindedAction>> KeyUpActions;
        private static bool _isDictionaryInitialised = false;

        private static IntPtr _currentSlideViewWindowHandle;


        private class KeyStatus
        {
            public bool IsPressed { get; private set; }
            public bool Ctrl { get; private set; }
            public bool Alt { get; private set; }
            public bool Shift { get; private set; }

            public void Press()
            {
                IsPressed = true;
                Ctrl = IsCtrlPressed();
                Shift = IsShiftPressed();
                Alt = IsAltPressed();
            }

            public void Release()
            {
                IsPressed = false;
            }
        }

        private struct BindedAction
        {
            private readonly bool Ctrl;
            private readonly bool Alt;
            private readonly bool Shift;
            private readonly Func<bool> ExecuteAction;

            public BindedAction(bool ctrl, bool alt, bool shift, Func<bool> action)
            {
                Ctrl = ctrl;
                Alt = alt;
                Shift = shift;
                ExecuteAction = action;
            }

            public bool RunConditionally(KeyStatus keyStatus)
            {
                if (Ctrl == keyStatus.Ctrl && Shift == keyStatus.Shift && Alt == keyStatus.Alt)
                {
                    return ExecuteAction();
                }
                return false;
            }
        }

        private static Native.HookProc hookProcedure;

        private static int hookId;
        private static bool Initialised = false;

        public static void Init(PowerPoint.Application application)
        {
            if (Initialised) return;
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

        private static void InitialiseDictionaries()
        {
            if (_isDictionaryInitialised) return;
            _isDictionaryInitialised = true;

            KeyStatuses = new Dictionary<int, KeyStatus>();
            KeyDownActions = new Dictionary<int, List<BindedAction>>();
            KeyUpActions = new Dictionary<int, List<BindedAction>>();

            foreach (var key in Enum.GetValues(typeof(Native.VirtualKey)))
            {
                int keyIndex = (int)key;
                KeyStatuses.Add(keyIndex, new KeyStatus());
                KeyDownActions.Add(keyIndex, new List<BindedAction>());
                KeyUpActions.Add(keyIndex, new List<BindedAction>());
            }
        }

        public static void CreateHook(IntPtr handle)
        {
            _currentSlideViewWindowHandle = FindSlideViewWindowHandle(handle);
            hookProcedure = HookProcedureCallback;
            hookId = Native.SetWindowsHookEx((int)Native.HookType.WH_KEYBOARD, hookProcedure, 0,
                Native.GetWindowThreadProcessId(_currentSlideViewWindowHandle, 0));
        }

        public static bool StopHook()
        {
            return Native.UnhookWindowsHookEx(hookId);
        }

        #region API
        public static void AddKeydownAction(Native.VirtualKey key, Action action, bool ctrl = false, bool alt = false, bool shift = false)
        {
            AddKeydownAction(key, ReturnFalse(action), ctrl, alt, shift);
        }

        public static void AddKeydownAction(Native.VirtualKey key, Func<bool> action, bool ctrl = false, bool alt = false, bool shift = false)
        {
            KeyDownActions[(int)key].Add(new BindedAction(ctrl, alt, shift, action));
        }

        public static void AddKeyupAction(Native.VirtualKey key, Action action, bool ctrl = false, bool alt = false, bool shift = false)
        {
            AddKeyupAction(key, ReturnFalse(action), ctrl, alt, shift);
        }

        public static void AddKeyupAction(Native.VirtualKey key, Func<bool> action, bool ctrl = false, bool alt = false, bool shift = false)
        {
            KeyUpActions[(int)key].Add(new BindedAction(ctrl, alt, shift, action));
        }

        public static void AddConditionToBlockTextInput(Func<bool> condition, bool ctrl = false, bool alt = false, bool shift = false)
        {
            Enum.GetValues(typeof (Native.VirtualKey)).Cast<Native.VirtualKey>()
                                                      .Where(Native.IsAlphanumericKey)
                                                      .ToList()
                                                      .ForEach(key => AddKeydownAction(key, condition, ctrl, alt, shift));
        }

        public static void AddConditionToBlockTextInput(Func<bool> condition, Native.VirtualKey key, bool ctrl = false, bool alt = false, bool shift = false)
        {
            AddKeydownAction(key, condition, ctrl, alt, shift);
        }
        #endregion


        /// <summary>
        /// A wrapper function for an Action that returns nothing, to make it into a Func&lt;bool&gt; that returns false.
        /// </summary>
        // A wrapper function for an Action that returns nothing, to make it into a Func<bool> that returns false.
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

        /// <summary>
        /// Returns true iff the main slide view window (the area which contains the slide) is focused by the user.
        /// </summary>
        private static bool IsSlideViewWindowFocused()
        {
            return Native.GetFocus() == _currentSlideViewWindowHandle;
        }

        private static int HookProcedureCallback(int nCode, IntPtr wParam, IntPtr lParam)
        {
            //Only process inputs that are sent to the main slide view window.
            if (!IsSlideViewWindowFocused()) return Native.CallNextHookEx(0, nCode, wParam, lParam);

            bool blockInput = false;
            if (nCode == 0)
            {
                int keyIndex = wParam.ToInt32();
                if (KeyStatuses.ContainsKey(keyIndex))
                {
                    var keyStatus = KeyStatuses[keyIndex];
                    if (IsKeydownCommand(lParam))
                    {
                        if (!keyStatus.IsPressed)
                        {
                            keyStatus.Press();
                            foreach (var action in KeyDownActions[keyIndex])
                            {
                                var block = action.RunConditionally(keyStatus);
                                if (block) blockInput = true;
                            }
                        }
                    }
                    else
                    {
                        if (keyStatus.IsPressed)
                        {
                            foreach (var action in KeyUpActions[keyIndex])
                            {
                                var block = action.RunConditionally(keyStatus);
                                if (block) blockInput = true;
                            }
                            keyStatus.Release();
                        }
                    }
                }
            }

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

        #region Modifier Keys
        private static bool IsCtrlPressed()
        {
            return IsModifierPressed(Native.VirtualKey.VK_LCONTROL) || IsModifierPressed(Native.VirtualKey.VK_RCONTROL);
        }

        private static bool IsAltPressed()
        {
            return IsModifierPressed(Native.VirtualKey.VK_LMENU) || IsModifierPressed(Native.VirtualKey.VK_RMENU);
        }

        private static bool IsShiftPressed()
        {
            return IsModifierPressed(Native.VirtualKey.VK_LSHIFT) || IsModifierPressed(Native.VirtualKey.VK_RSHIFT);
        }

        /// <summary>
        /// Used to check whether the Ctrl, Alt or Shift keys are being held down.
        /// </summary>
        private static bool IsModifierPressed(Native.VirtualKey key)
        {
            return (Native.GetKeyState(key) & 0x80000000) != 0;
        }
        #endregion

        private static bool IsHooked()
        {
            return hookId != 0;
        }

    }
}