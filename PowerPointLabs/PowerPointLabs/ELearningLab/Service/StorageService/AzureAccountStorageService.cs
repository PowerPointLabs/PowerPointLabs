using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Xml.Serialization;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.AudioGenerator;

namespace PowerPointLabs.ELearningLab.Service
{
    public static class AzureAccountStorageService
    {
        private const string defaultNarrationsStorageFolder = "PowerPointLabs Narrations Access Key Storage";
        private const string defaultNarrationsStorageFile = "useraccount.xml";
        private static string defaultApplicationFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        public static string GetELearningLabStoragePath()
        {
            return Path.Combine(defaultApplicationFolderPath, defaultNarrationsStorageFolder);
        }

        public static string GetAccessKeyFilePath()
        {
            return Path.Combine(defaultApplicationFolderPath, defaultNarrationsStorageFolder, defaultNarrationsStorageFile);
        }

        public static void SaveUserAccount(AzureAccount account)
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
            if (!AzureAccount.GetInstance().IsEmpty())
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
                string endpoint = user.ContainsKey("endpoint") ? user["endpoint"].Trim() : null;
                if (key != null && endpoint != null)
                {
                    AzureAccount.GetInstance().SetUserKeyAndRegion(key, endpoint);
                    AzureRuntimeService.IsAzureAccountPresentAndValid = 
                        AzureRuntimeService.IsValidUserAccount(errorMessage: "Invalid Azure Account." +
                        "\nIs your Azure account expired?\nAre you connected to Wifi?");
                }
                else
                {
                    AzureRuntimeService.IsAzureAccountPresentAndValid = false;
                    File.Delete(GetAccessKeyFilePath());
                }
            }
            catch (Exception e)
            {
                AzureRuntimeService.IsAzureAccountPresentAndValid = false;
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
