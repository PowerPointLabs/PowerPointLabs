using System;
using System.Collections.Generic;
using System.Windows;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    public class PPLClipboard
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetOpenClipboardWindow();

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern bool CloseClipboard();

        public static PPLClipboard Instance
        {
            get
            {
                if (_instance == null)
                {
                    _instance = new PPLClipboard();
                }
                return _instance;
            }
        }

        private static PPLClipboard _instance;
        private bool isLocked = false;
        private Dictionary<string, object> lBackup = new Dictionary<string, object>();
        private IDataObject lDataObject;
        private string[] lFormats;

        public void LockClipboard()
        {
            if (isLocked) { return; }
            // wait to lock the clipboard
            while (GetOpenClipboardWindow() != IntPtr.Zero)
            {
                MessageBox.Show("Another application is currently using the clipboard. Please come back later and try again", "Retry", MessageBoxButton.OK);
            }
            OpenClipboard(new IntPtr(Globals.ThisAddIn.Application.ActiveWindow.HWND));
            SaveClipboard();
            isLocked = true;
        }

        public void ReleaseClipboard()
        {
            if (!isLocked) { return; }
            RestoreClipboard();
            CloseClipboard();
            isLocked = false;
        }

        private void SaveClipboard()
        {
            lDataObject = Clipboard.GetDataObject();
            lFormats = lDataObject.GetFormats(false);
            foreach (var lFormat in lFormats)
            {
                lBackup.Add(lFormat, lDataObject.GetData(lFormat, false));
            }
            Clipboard.Clear(); // to prevent any weird data from showing up
        }

        // referred to https://stackoverflow.com/questions/6262454/c-sharp-backing-up-and-restoring-clipboard
        private void RestoreClipboard()
        {
            foreach (var lFormat in lFormats)
            {
                lDataObject.SetData(lBackup[lFormat]);
            }
            Clipboard.SetDataObject(lDataObject);
            Clipboard.Flush();
            lBackup.Clear();
            lDataObject = null;
            lFormats = null;
        }

        public void Teardown()
        {
            ReleaseClipboard();
            _instance = null;
        }
    }
}
