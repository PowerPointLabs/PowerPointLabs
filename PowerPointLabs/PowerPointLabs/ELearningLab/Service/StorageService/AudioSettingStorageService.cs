using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml.Linq;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.AudioGenerator;

namespace PowerPointLabs.ELearningLab.Service
{
    public class AudioSettingStorageService
    {
        private const string defaultELearningLabStorageFolder = "PowerPointLabs Narrations Access Key Storage";
        private const string defaultAudioSettingStorageFile = "audioSettingPreference.xml";
        private static string defaultApplicationFolderPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

        public static string GetELearningLabStoragePath()
        {
            return Path.Combine(defaultApplicationFolderPath, defaultELearningLabStorageFolder);
        }

        public static string GetAudioSettingFilePath()
        {
            return Path.Combine(defaultApplicationFolderPath, defaultELearningLabStorageFolder, defaultAudioSettingStorageFile);
        }

        public static void SaveAudioSettingPreference()
        {
            if (!Directory.Exists(GetELearningLabStoragePath()))
            {
                Directory.CreateDirectory(GetELearningLabStoragePath());
            }
            if (File.Exists(GetAudioSettingFilePath()))
            {
                try
                {
                    File.Delete(GetAudioSettingFilePath());
                }
                catch
                {
                    Logger.Log("Cannot delete audio setting files because other presentations are using it.");
                }
            }
            Dictionary<string, string> audioSetting = new Dictionary<string, string>()
            {
                { "voiceName", AudioSettingService.selectedVoice.ToString() }
            };
            XElement root = new XElement("audioSetting", from kv in audioSetting select new XElement(kv.Key, kv.Value));
            FileStream file = File.Open(GetAudioSettingFilePath(), FileMode.OpenOrCreate);
            root.Save(file);
            file.Close();
        }

        public static void LoadAudioSettingPreference()
        {
            Dictionary<string, string> audioSetting = new Dictionary<string, string>();
            try
            {
                FileStream file = File.Open(GetAudioSettingFilePath(), FileMode.Open);
                XElement root = XElement.Load(file);
                foreach (XElement el in root.Elements())
                {
                    audioSetting.Add(el.Name.LocalName, el.Value);
                }
                string voiceName = audioSetting.ContainsKey("voiceName") ? audioSetting["voiceName"] : null;
                if (voiceName != null)
                {
                    IVoice voice = AudioService.GetVoiceFromString(voiceName.Trim());
                    AudioSettingService.selectedVoice = voice;
                    if (voice is ComputerVoice)
                    {
                        AudioSettingService.selectedVoiceType = VoiceType.ComputerVoice;
                    }
                    else if (voice is AzureVoice)
                    {
                        AudioSettingService.selectedVoiceType = VoiceType.AzureVoice;
                    }
                    else
                    {
                        Logger.Log("Error: Voice retrieved has invalid type");
                    }
                }
                else
                {
                    File.Delete(GetAudioSettingFilePath());
                }
            }
            catch (Exception e)
            {
                // handle exception
                Logger.Log(e.Message);
            }
        }

        public static void DeleteAudioSettingPreference()
        {
            try
            {
                if (Directory.Exists(GetELearningLabStoragePath()) && File.Exists(GetAudioSettingFilePath()))
                {
                    File.Delete(GetAudioSettingFilePath());
                }
            }
            catch (Exception e)
            {
                Logger.Log(e.Message);
            }
        }
    }
}
