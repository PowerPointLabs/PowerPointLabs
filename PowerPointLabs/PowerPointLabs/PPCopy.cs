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

        public static void Init(PowerPoint.Application application)
        {
            if (instance == null)
            {
                instance = new PPCopy();

                instance.Visible = false;
                isSuccessful = Native.AddClipboardFormatListener(instance.Handle);
                application.WindowSelectionChange += (selection) =>
                {
                    selectedRange = selection;
                    if (!isSuccessful)
                    {
                        isSuccessful = Native.AddClipboardFormatListener(instance.Handle);
                    }
                };
            }
        }

        //Delegate
        public delegate void CopyEventDelegate(PowerPoint.Selection selection);

        //Handler
        public static event CopyEventDelegate AfterCopyPaste;

        protected override void Dispose(bool disposing)
        {
            Native.RemoveClipboardFormatListener(instance.Handle);
        }

        protected override void WndProc(ref Message m)
        {
            base.WndProc(ref m);
            if (m.Msg == (int)Native.Message.WM_CLIPBOARDUPDATE)
            {
                if (selectedRange != null
                    && AfterCopyPaste != null)
                {
                    PPCopy.AfterCopyPaste(selectedRange);
                }
            }
        }
    }
}
