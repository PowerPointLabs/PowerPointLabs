using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Serialization;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.NarrationsLab.Data;

namespace PowerPointLabs.NarrationsLab.Storage
{
    public static class NarrationsLabStorageConfig
    {
        private const string defaultNarrationsStorageFolder = "PowerPointLabs Narrations Access Key Storage";
        private const string defaultNarrationsStorageFile = "useraccount.xml";
        private static string defaultApplicationFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        public static string GetAccessKeyStoragePath()
        {
            return Path.Combine(defaultApplicationFolderPath, defaultNarrationsStorageFolder);
        }

        public static string GetAccessKeyFilePath()
        {
            return Path.Combine(defaultApplicationFolderPath, defaultNarrationsStorageFolder, defaultNarrationsStorageFile);
        }

        public static void SaveUserAccount(UserAccount account)
        {
            if (!Directory.Exists(GetAccessKeyStoragePath()))
            {
                Directory.CreateDirectory(GetAccessKeyStoragePath());
            }
          
            Dictionary<string, string> user = new Dictionary<string, string>()
            {
                { "endpoint", account.GetRegion() },
                {"key",  account.GetKey() }
            };
            XElement root = new XElement("user", from kv in user select new XElement(kv.Key, kv.Value));
            FileStream file = File.Open(GetAccessKeyFilePath(), FileMode.OpenOrCreate);
            root.Save(file);
            file.Close();
        }

        public static void LoadUserAccount()
        {
            Dictionary<string, string> user = new Dictionary<string, string>();
            try
            {
                FileStream file = File.Open(GetAccessKeyFilePath(), FileMode.Open);
                XElement root = XElement.Load(file);
                foreach (XElement el in root.Elements())
                {
                    user.Add(el.Name.LocalName, el.Value);
                }
                string key = user.ContainsKey("key") ? user["key"] : null;
                string endpoint = user.ContainsKey("endpoint") ? user["endpoint"] : null;
                if (key != null && endpoint != null)
                {
                    UserAccount.GetInstance().SetUserKeyAndRegion(key, endpoint);
                }
                else
                {
                    File.Delete(GetAccessKeyFilePath());
                }
            }
            catch (Exception e)
            {
                // handle exception
                Logger.Log(e.Message);
            }
        }

        public static void DeleteUserAccount()
        {
            if (Directory.Exists(GetAccessKeyStoragePath()) && File.Exists(GetAccessKeyFilePath()))
            {
                File.Delete(GetAccessKeyFilePath());
            }
        }
    }
}
