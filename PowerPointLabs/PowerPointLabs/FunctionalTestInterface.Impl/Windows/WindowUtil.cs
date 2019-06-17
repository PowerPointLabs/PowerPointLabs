﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.InteropServices;
using System.Text;
using static PPExtraEventHelper.Native;
using HWND = System.IntPtr;

namespace PowerPointLabs.FunctionalTestInterface.Windows
{
    /// <summary>
    /// An utility class for retrieving windows and handles;
    /// </summary>
    public static class WindowUtil
    {
        private const uint WM_CLOSE = 0x0010;

        public static void CloseWindow(HWND hwnd)
        {
            SendMessage(hwnd, WM_CLOSE, HWND.Zero, HWND.Zero);
        }

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

        internal static HWND SubscribeActiveWindowChanged(WinEventDelegate callback)
        {
            HWND windowEventHook = SetWinEventHook(EVENT_SYSTEM_FOREGROUND,
                EVENT_SYSTEM_FOREGROUND, HWND.Zero,
                callback, 0, 0,
                WINEVENT_OUTOFCONTEXT | WINEVENT_SKIPOWNPROCESS);
            if (windowEventHook == HWND.Zero)
            {
                throw new Win32Exception(Marshal.GetLastWin32Error());
            }
            return windowEventHook;
        }

        internal static void UnsubscribeActiveWindowChanged(HWND hwnd)
        {
            UnhookWinEvent(hwnd);
        }

        private const int WINEVENT_INCONTEXT = 4;
        private const int WINEVENT_OUTOFCONTEXT = 0;
        private const int WINEVENT_SKIPOWNPROCESS = 2;
        private const int WINEVENT_SKIPOWNTHREAD = 1;

        private const int EVENT_SYSTEM_FOREGROUND = 3;

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
