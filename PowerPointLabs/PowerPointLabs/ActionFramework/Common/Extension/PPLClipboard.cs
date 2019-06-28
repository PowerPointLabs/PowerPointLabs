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
        public bool IsLocked { get; private set; }
        private Dictionary<string, object> lBackup = new Dictionary<string, object>();
        private IDataObject lDataObject;
        private IntPtr _parentWindow => new IntPtr(Globals.ThisAddIn.Application.HWND);
        public bool AutoDismiss = false;

        private PPLClipboard()
        {
            IsLocked = false;
        }

        public void LockAndRelease(Action action)
        {
            LockAndRelease<object>(() =>
            {
                action();
                return null;
            });
        }

        public TResult LockAndRelease<TResult>(Func<TResult> action)
        {
            LockClipboard();
            try
            {
                return action();
            }
            catch (Exception e)
            {
                throw e;
            }
            finally
            {
                ReleaseClipboard();
            }
        }

        public void LockClipboard()
        {
            if (IsLocked) { throw new Exception("Clipboard is not released before locking!"); }
            // wait to lock the clipboard
            while (!IsClipboardFree() && !OpenClipboard(_parentWindow))
            {
                if (!AutoDismiss)
                {
                    MessageBox.Show("Another application is currently using the clipboard. Please come back later and try again", "Retry", MessageBoxButton.OK);
                }
            }
            // PowerPointShapeGalleryPresentation has a strong reliance on the clipboard
            SaveClipboard();
            IsLocked = true;
        }

        public void ReleaseClipboard()
        {
            if (!IsLocked) { return; }
            RestoreClipboard();
            CloseClipboard();
            IsLocked = false;
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
            if (lDataObject != null)
            {
                Clipboard.SetDataObject(lDataObject);
            }
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
