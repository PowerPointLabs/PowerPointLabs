using System;
using System.Windows.Forms;

using PPExtraEventHelper;

namespace PowerPointLabs.Views
{
    // this class is taken from
    // http://stackoverflow.com/questions/6942150/why-folderbrowserdialog-dialog-does-not-scroll-to-selected-folder
    // with a bit refactoring.
    static class FolderDialogLauncher
    {
        /// <summary>
        /// Using title text to look for the top level dialog window is fragile.
        /// In particular, this will fail in non-English applications.
        /// </summary>
        private const string TopLevelSearchString = "Browse For Folder";

        /// <summary>
        /// These should be more robust.  We find the correct child controls in the dialog
        /// by using the GetDlgItem method, rather than the FindWindow(Ex) method,
        /// because the dialog item IDs should be constant.
        /// </summary>
        private const int DlgItemBrowseControl = 0;
        private const int DlgItemTreeView = 100;

        private static int _retries = 10;

        /// <summary>
        /// Calling this method is identical to calling the ShowDialog method of the provided
        /// FolderBrowserDialog, except that an attempt will be made to scroll the Tree View
        /// to make the currently selected folder visible in the dialog window.
        /// </summary>
        /// <param name="dlg"></param>
        /// <param name="parent"></param>
        /// <returns></returns>
        public static DialogResult ShowFolderBrowser(FolderBrowserDialog dlg, IWin32Window parent = null)
        {
            DialogResult result;

            using (Timer timer = new Timer())
            {
                timer.Tick += TimerTickHandler;
                timer.Interval = 10;
                timer.Start();

                result = dlg.ShowDialog(parent);
            }

            _retries = 10;
            return result;
        }

        private static void TimerTickHandler(object sender, EventArgs args)
        {
            Timer timer = sender as Timer;

            if (timer == null)
            {
                return;
            }

            if (_retries > 0)
            {
                // retry 10 times until the dialog handle is created since we have no
                // means to override wndproc of FolderBrowswerDialog to find handle.
                --_retries;
                IntPtr hwndDlg = Native.FindWindow(null, TopLevelSearchString);
                if (hwndDlg != IntPtr.Zero)
                {
                    IntPtr hwndFolderCtrl = Native.GetDlgItem(hwndDlg, DlgItemBrowseControl);
                    if (hwndFolderCtrl != IntPtr.Zero)
                    {
                        IntPtr hwndTreeView = Native.GetDlgItem(hwndFolderCtrl, DlgItemTreeView);

                        if (hwndTreeView != IntPtr.Zero)
                        {
                            IntPtr item = Native.SendMessage(hwndTreeView, (uint)Native.Message.TVM_GETNEXTITEM,
                                                          new IntPtr((uint)Native.Message.TVGN_CARET),
                                                          IntPtr.Zero);
                            if (item != IntPtr.Zero)
                            {
                                Native.SendMessage(hwndTreeView, (uint)Native.Message.TVM_ENSUREVISIBLE, IntPtr.Zero, item);
                                _retries = 0;
                                timer.Stop();
                            }
                        }
                    }
                }
            }
            else
            {
                //  We failed to find the Tree View control window.
                //
                //  As a fall back (and this is an UberUgly hack), we will send
                //  some fake keystrokes to the application in an attempt to force
                //  the Tree View to scroll to the selected item.
                //
                //  This method may not always work. On some laggy machine or virtual
                //  machine, key strokes may not reach the dialog successfully, either
                //  partly or entirely lost.
                timer.Stop();
                SendKeys.Send("{TAB}{TAB}{DOWN}{UP}");
            }
        }
    }
}
