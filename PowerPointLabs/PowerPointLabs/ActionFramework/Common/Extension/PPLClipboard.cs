using System;
using System.Collections.Generic;
using System.Windows;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    public class PPLClipboard
    {
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
        private IntPtr _parentWindow => new IntPtr(Globals.ThisAddIn.Application.HWND);

        public void LockClipboard()
        {
            if (isLocked) { return; }
            // wait to lock the clipboard
            while (!IsClipboardFree() && !OpenClipboard(_parentWindow))
            {
                MessageBox.Show("Another application is currently using the clipboard. Please come back later and try again", "Retry", MessageBoxButton.OK);
            }
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

        public void Teardown()
        {
            ReleaseClipboard();
            _instance = null;
        }

        private bool IsClipboardFree()
        {
            IntPtr hwnd = GetOpenClipboardWindow();
            return hwnd == IntPtr.Zero || hwnd == _parentWindow;
        }

        private void SaveClipboard()
        {
            lDataObject = Clipboard.GetDataObject();
            string[] lFormats = lDataObject.GetFormats(false);
            lBackup.Clear();
            /*
            foreach (string lFormat in lFormats)
            {
                if (lDataObject.GetDataPresent(lFormat))
                {
                    lBackup.Add(lFormat, lDataObject.GetData(lFormat, false));
                }
            }
            */
            //Clipboard.Clear(); // to prevent any weird data from showing up
        }

        // referred to https://stackoverflow.com/questions/6262454/c-sharp-backing-up-and-restoring-clipboard
        private void RestoreClipboard()
        {
            /*
            foreach (KeyValuePair<string, object> pair in lBackup)
            {
                lDataObject.SetData(pair.Key, pair.Value, false);
            }
            */
            Clipboard.SetDataObject(lDataObject);
            Clipboard.Flush();
            lBackup.Clear();
            lDataObject = null;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetOpenClipboardWindow();

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [System.Runtime.InteropServices.DllImport("user32.dll", SetLastError = true)]
        private static extern bool CloseClipboard();

    }
}
