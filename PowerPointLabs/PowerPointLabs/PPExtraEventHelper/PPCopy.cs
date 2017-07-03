using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PPExtraEventHelper
{
    internal class PPCopy : Control
    {
        private static bool isSuccessful;
        private static PowerPoint.Selection selectedRange;
        private static PPCopy instance;
        private static IntPtr handle;
        private static IntPtr nextHandle;
        private static bool isFirstTimeEnter = true;
        private static bool isCopyEvent = false;
        private static bool isDisposed = false;

        public static void Init(PowerPoint.Application application)
        {
            if (instance == null)
            {
                instance = new PPCopy();
                instance.Visible = false;
                try
                {
                    nextHandle = Native.SetClipboardViewer(instance.Handle);
                    isSuccessful = Native.AddClipboardFormatListener(instance.Handle);
                    //if fails to set up clipboardFormatListener, re-try 5 times
                    for (int i = 0; i < 5; i++)
                    {
                        if (isSuccessful)
                        {
                            break;
                        }

                        isSuccessful = Native.AddClipboardFormatListener(instance.Handle);
                    }

                    handle = instance.Handle;
                    application.WindowSelectionChange += (selection) =>
                    {
                        selectedRange = selection;
                        if (!isSuccessful)
                        {
                            isSuccessful = Native.AddClipboardFormatListener(instance.Handle);
                        }
                    };
                }
                catch
                {
                    //TODO: support win XP?
                }
            }
        }

        //Delegate
        public delegate void CopyEventDelegate(PowerPoint.Selection selection);

        //Handler
        public static event CopyEventDelegate AfterCopy;

        public static event CopyEventDelegate AfterPaste;

        public static void StopHook()
        {
            if (isDisposed)
            {
                return;
            }

            isDisposed = true;
            instance.Dispose(true);
        }

        protected override void Dispose(bool disposing)
        {
            Native.ChangeClipboardChain(handle, nextHandle);
            if (isSuccessful)
            {
                Native.RemoveClipboardFormatListener(handle);
            }
            base.Dispose(disposing);
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == (int)Native.Message.WM_CHANGECBCHAIN)
            {
                if (m.WParam == nextHandle)
                {
                    nextHandle = m.LParam;
                }
                else
                {
                    Native.SendMessage(nextHandle, (uint)m.Msg, m.WParam, m.LParam);
                }
            }
            else if (m.Msg == (int)Native.Message.WM_DRAWCLIPBOARD)
            {
                if (!isFirstTimeEnter)
                {
                    isCopyEvent = true;
                }
                else
                {
                    isFirstTimeEnter = false;
                }
                Native.SendMessage(nextHandle, (uint)m.Msg, m.WParam, m.LParam);
            }
            else if (m.Msg == (int)Native.Message.WM_CLIPBOARDUPDATE)
            {
                try
                {
                    if (selectedRange != null
                        && selectedRange.Type != PowerPoint.PpSelectionType.ppSelectionNone
                        && AfterCopy != null
                        && AfterPaste != null)
                    {
                        if (isCopyEvent)
                        {
                            AfterCopy(selectedRange);
                        }
                        else //PasteEvent
                        {
                            AfterPaste(selectedRange);
                        }
                        isCopyEvent = false;
                    }
                }
                catch (Exception)
                {
                    
                }
            }
        }
    }
}
