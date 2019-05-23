using System;
using System.Management;
using System.Security.Principal;
using Microsoft.Win32;

namespace PowerPointLabs.Utils
{
    /// <summary>
    /// A class that allows watching of Registry Key values.
    /// </summary>
    class RegistryWatcher<T>
    {
        private readonly string path;
        private readonly string key;
        private ManagementEventWatcher watcher;

        public event EventHandler<T> ValueChanged;

        public RegistryWatcher(string path, string key)
        {
            this.path = path;
            this.key = key;
            RegisterKeyChanged();
        }
 
        /// <summary>
        /// Fires the event manually
        /// </summary>
        public void Fire()
        {
            Notify();
        }

        public void Start()
        {
            watcher.Start();
        }
        public void Stop()
        {
            watcher.Stop();
        }

        public void SetValue(object o)
        {
            WindowsIdentity currentUser = WindowsIdentity.GetCurrent();
            Registry.SetValue(String.Format("{0}\\{1}", currentUser.User.Value, path), key, o);
        }

        private void RegisterKeyChanged()
        {
            WindowsIdentity currentUser = WindowsIdentity.GetCurrent();
            WqlEventQuery query = new WqlEventQuery(
                     "SELECT * FROM RegistryValueChangeEvent WHERE " +
                     "Hive = 'HKEY_USERS'" +
                String.Format(@"AND KeyPath = '{0}\\{1}' AND ValueName='{2}'", currentUser.User.Value, path, key));
            watcher = new ManagementEventWatcher(query);
            watcher.EventArrived += (object sender, EventArrivedEventArgs e) => { Notify(); };
        }

        private T GetKey()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(path))
            {
                object objectValue;
                if (key == null || (objectValue = key.GetValue(this.key)) == null)
                {
                    throw new Exceptions.AssumptionFailedException("Key is null");
                }
                return (T)objectValue;
            }
        }

        private void Notify()
        {
            try
            {
                T key = GetKey();
                if (ValueChanged != null)
                {
                    ValueChanged(this, key);
                }
            }
            catch (Exception)
            {

            }
        }
    }
}
