using System;
using System.Collections.Generic;
using System.Collections.Specialized;
using System.Windows;

namespace PowerPointLabs.ActionFramework.Common.Extension
{
    public class PPLClipboard
    {
        public static PPLClipboard Instance { get; private set; }
        public bool IsLocked { get; private set; }
        private IDataObject lDataObject;
        private IntPtr _parentWindow;
        public bool AutoDismiss;

        private PPLClipboard(IntPtr parentWindow, bool autoDismiss)
        {
            IsLocked = false;
            _parentWindow = parentWindow;
            AutoDismiss = autoDismiss;
        }

        public bool IsEmpty()
        {
            return LockIfNeeded(() =>
            {
                IDataObject clipboardData = Clipboard.GetDataObject();
                return clipboardData == null || clipboardData.GetFormats().Length == 0;
            });
        }

        public System.Drawing.Image GetImage()
        {
            return LockIfNeeded(() =>
            {
                return System.Windows.Forms.Clipboard.GetImage();
            });
        }

        public StringCollection GetFileDropList()
        {
            return LockIfNeeded(() =>
            {
                return Clipboard.GetFileDropList();
            });
        }

        public List<object> LoadClipboardObjects()
        {
            return LockIfNeeded(() =>
            {
                List<object> result = new List<object>();
                if (Clipboard.ContainsImage())
                {
                    result.Add(Clipboard.GetImage());
                }
                if (Clipboard.ContainsFileDropList())
                {
                    result.Add(Clipboard.GetFileDropList());
                }
                if (Clipboard.ContainsText())
                {
                    result.Add(Clipboard.GetText());
                }
                return result;
            });
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
            if (IsLocked) {
                throw new Exception("Clipboard is not released before locking!");
            }
            // wait to lock the clipboard
            while (!IsClipboardFree() && !OpenClipboard(_parentWindow))
            {
                if (!AutoDismiss)
                {
                    MessageBox.Show("Another application is currently using the clipboard. Please come back later and try again", "Retry", MessageBoxButton.OK);
                }
            }
            IsLocked = true;
        }

        public void ReleaseClipboard()
        {
            if (!IsLocked) { return; }
            CloseClipboard();
            IsLocked = false;
        }

        public static void Init(IntPtr parentWindow, bool autoDismiss = false)
        {
            if (Instance == null)
            {
                Instance = new PPLClipboard(parentWindow, autoDismiss);
            }
        }

        public void Teardown()
        {
            ReleaseClipboard();
            Instance = null;
        }

        // not working stably
        public void SaveClipboard()
        {
            lDataObject = Clipboard.GetDataObject();
        }

        // referred to https://stackoverflow.com/questions/6262454/c-sharp-backing-up-and-restoring-clipboard
        // not working stably
        public void RestoreClipboard()
        {
            if (lDataObject != null)
            {
                Clipboard.SetDataObject(lDataObject);
            }
            Clipboard.Flush();
            lDataObject = null;
        }

        private TResult LockIfNeeded<TResult>(Func<TResult> action)
        {
            if (IsLocked)
            {
                return action();
            }
            LockClipboard();
            TResult result = action();
            ReleaseClipboard();
            return result;
        }

        private bool IsClipboardFree()
        {
            IntPtr hwnd = GetOpenClipboardWindow();
            return hwnd == IntPtr.Zero || hwnd == _parentWindow;
        }

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern IntPtr GetOpenClipboardWindow();

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool OpenClipboard(IntPtr hWndNewOwner);

        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern bool CloseClipboard();

    }
}
