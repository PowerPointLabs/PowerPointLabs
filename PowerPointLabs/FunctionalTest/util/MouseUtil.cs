using System;
using System.Drawing;
using System.Windows.Forms;

namespace FunctionalTest.util
{
    /// <summary>
    /// ref: http://stackoverflow.com/questions/8739523/directing-mouse-events-dllimportuser32-dll-click-double-click
    /// </summary>
    class MouseUtil
    {
        public static void SendMouseLeftClick(int x, int y)
        {
            Cursor.Position = new Point(x, y);
            NativeUtil.mouse_event(
                NativeUtil.MOUSEEVENTF_LEFTDOWN | NativeUtil.MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
        }

        public static void SendMouseRightclick(int x, int y)
        {
            Cursor.Position = new Point(x, y);
            NativeUtil.mouse_event(
                NativeUtil.MOUSEEVENTF_RIGHTDOWN | NativeUtil.MOUSEEVENTF_RIGHTUP, 0, 0, 0, UIntPtr.Zero);
        }

        public static void SendMouseDoubleClick(int x, int y)
        {
            Cursor.Position = new Point(x, y);
            NativeUtil.mouse_event(
                NativeUtil.MOUSEEVENTF_LEFTDOWN | NativeUtil.MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);

            ThreadUtil.WaitFor(150);

            NativeUtil.mouse_event(
                NativeUtil.MOUSEEVENTF_LEFTDOWN | NativeUtil.MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
        }

        public static void SendMouseRightDoubleClick(int x, int y)
        {
            Cursor.Position = new Point(x, y);
            NativeUtil.mouse_event(
                NativeUtil.MOUSEEVENTF_RIGHTDOWN | NativeUtil.MOUSEEVENTF_RIGHTUP, 0, 0, 0, UIntPtr.Zero);

            ThreadUtil.WaitFor(150);

            NativeUtil.mouse_event(
                NativeUtil.MOUSEEVENTF_RIGHTDOWN | NativeUtil.MOUSEEVENTF_RIGHTUP, 0, 0, 0, UIntPtr.Zero);
        }

        public static void SendMouseDown(int x, int y)
        {
            Cursor.Position = new Point(x, y);
            NativeUtil.mouse_event(NativeUtil.MOUSEEVENTF_LEFTDOWN, 0, 0, 0, UIntPtr.Zero);
            ThreadUtil.WaitFor(1000);
        }

        public static void SendMouseUp(int x, int y)
        {
            Cursor.Position = new Point(x, y);
            ThreadUtil.WaitFor(1000);
            NativeUtil.mouse_event(NativeUtil.MOUSEEVENTF_LEFTUP, 0, 0, 0, UIntPtr.Zero);
            ThreadUtil.WaitFor(150);
        }
    }
}
