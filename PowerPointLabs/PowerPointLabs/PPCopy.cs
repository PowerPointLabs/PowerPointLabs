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
                    handle = instance.Handle;
                    application.WindowSelectionChange += (selection) =>
                    {
                        selectedRange = selection;
                        if (!isSuccessful)
                        {
                            nextHandle = Native.SetClipboardViewer(instance.Handle);
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

        protected override void Dispose(bool disposing)
        {
            if (isSuccessful)
            {
                Native.ChangeClipboardChain(handle, nextHandle);
                Native.RemoveClipboardFormatListener(handle);
            }
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
            }
            else if (m.Msg == (int)Native.Message.WM_CLIPBOARDUPDATE)
            {
                if (selectedRange != null
                    && selectedRange.Type != PowerPoint.PpSelectionType.ppSelectionNone
                    && AfterCopy != null
                    && AfterPaste != null)
                {
                    if (isCopyEvent)
                    {
                        PREVENT_FINITE_LOOP_BEGIN();
                        PPCopy.AfterCopy(selectedRange);
                        PREVENT_FINITE_LOOP_END();
                    }
                    else//PasteEvent
                    {
                        PREVENT_FINITE_LOOP_BEGIN();
                        PPCopy.AfterPaste(selectedRange);
                        PREVENT_FINITE_LOOP_END();
                    }
                    isCopyEvent = false;
                }
            }
        }

        private static void PREVENT_FINITE_LOOP_BEGIN()
        {
            Native.RemoveClipboardFormatListener(handle);
        }

        private static void PREVENT_FINITE_LOOP_END()
        {
            isSuccessful = Native.AddClipboardFormatListener(instance.Handle);
        }
    }
}
