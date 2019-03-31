using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Xml.Linq;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.AudioGenerator.WatsonVoiceGenerator.Model;

namespace PowerPointLabs.ELearningLab.Service.StorageService
{
    public static class WatsonAccountStorageService
    {
        private const string defaultNarrationsStorageFolder = "PowerPointLabs Narrations Access Key Storage";
        private const string defaultNarrationsStorageFile = "watsonaccount.xml";
        private static string defaultApplicationFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        public static string GetELearningLabStoragePath()
        {
            return Path.Combine(defaultApplicationFolderPath, defaultNarrationsStorageFolder);
        }

        public static string GetAccessKeyFilePath()
        {
            return Path.Combine(defaultApplicationFolderPath, defaultNarrationsStorageFolder, defaultNarrationsStorageFile);
        }

        public static void SaveUserAccount(WatsonAccount account)
        {
            if (!Directory.Exists(GetELearningLabStoragePath()))
            {
                Directory.CreateDirectory(GetELearningLabStoragePath());
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
            if (!WatsonAccount.GetInstance().IsEmpty())
            {
                return;
            }
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
                    WatsonAccount.GetInstance().SetUserKeyAndRegion(key, endpoint);
                    WatsonRuntimeService.IsWatsonAccountPresentAndValid = WatsonRuntimeService.IsValidUserAccount();
                }
                else
                {
                    WatsonRuntimeService.IsWatsonAccountPresentAndValid = false;
                    File.Delete(GetAccessKeyFilePath());
                }
            }
            catch (Exception e)
            {
                WatsonRuntimeService.IsWatsonAccountPresentAndValid = false;
                Logger.Log(e.Message);
            }
        }

        public static void DeleteUserAccount()
        {
            try
            {
                if (Directory.Exists(GetELearningLabStoragePath()) && File.Exists(GetAccessKeyFilePath()))
                {
                    File.Delete(GetAccessKeyFilePath());
                }
            }
            catch (Exception e)
            {
                Logger.Log(e.Message);
            }
        }
    }
}
