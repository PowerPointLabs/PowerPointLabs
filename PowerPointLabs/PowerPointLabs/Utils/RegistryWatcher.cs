using System;
using System.Collections.Generic;
using System.Management;
using System.Security.Principal;
using Microsoft.Win32;
using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.Utils
{
    /// <summary>
    /// A class that allows watching of Registry Key values.
    /// </summary>
    class RegistryWatcher<T> where T : IEquatable<T>
    {
        private readonly string path;
        private readonly string key;
        private readonly List<T> defaultKey;
        private ManagementEventWatcher watcher;

        // returns true if the key started as defaultKey and is not modified, else false
        public bool IsDefaultKey { get; private set; }

        public event EventHandler<T> ValueChanged;

        public RegistryWatcher(string path, string key, List<T> defaultKey)
        {
            this.path = path;
            this.key = key;
            this.defaultKey = defaultKey;
            this.IsDefaultKey = true;
            RegisterKeyChanged();
            GetKeyAndUpdateKeyStatus();
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

        private T GetKeyAndUpdateKeyStatus()
        {
            using (RegistryKey key = Registry.CurrentUser.OpenSubKey(path))
            {
                object objectValue;
                if (key == null || (objectValue = key.GetValue(this.key)) == null)
                {
                    throw new Exceptions.AssumptionFailedException("Key is null");
                }
                T result = (T)objectValue;
                IsDefaultKey &= defaultKey == null || defaultKey.Contains(result);
                return result;
            }
        }

        private void Notify()
        {
            try
            {
                T key = GetKeyAndUpdateKeyStatus();
                if (IsDefaultKey)
                {
                    return;
                }
                ValueChanged?.Invoke(this, key);
            }
            catch (Exception e)
            {
                Logger.LogException(e, nameof(Notify));
            }
        }
    }
}
