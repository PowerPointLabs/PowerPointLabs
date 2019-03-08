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
            List<string> audios = ConvertAudioPreferencesToList();
            XElement root = new XElement("audioSetting", 
                audios.Select( x =>
                new XElement("audio", new XAttribute("name", x))));
            using (FileStream file = File.Open(GetAudioSettingFilePath(), FileMode.OpenOrCreate))
            {
                root.Save(file);
            }
        }

        public static void LoadAudioSettingPreference()
        {
            List<string> audioPreference = new List<string>();
            try
            {
                using (FileStream file = File.Open(GetAudioSettingFilePath(), FileMode.Open))
                {
                    XElement root = XElement.Load(file);
                    foreach (XElement el in root.Elements())
                    {
                        audioPreference.Add(el.Attribute("name").Value);
                    }
                    if (audioPreference.Count == 0)
                    {
                        File.Delete(GetAudioSettingFilePath());
                        return;
                    }
                    InitializeDefaultAndRankedVoicesFromList(audioPreference);
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

        private static List<string> ConvertAudioPreferencesToList()
        {
            IVoice defaultVoice = AudioSettingService.selectedVoice;
            List<string> voices = AudioSettingService.preferredVoices.ToList().Select(x => x.VoiceName).ToList();
            voices.Insert(0, defaultVoice.ToString());
            return voices;
        }

        private static void InitializeDefaultAndRankedVoicesFromList(List<string> preferences)
        {
            for (int i = 0;  i < preferences.Count(); i++)
            {
                string voiceName = preferences[i].Trim();
                IVoice voice = AudioService.GetVoiceFromString(voiceName);
                voice.Rank = i;
                if (i == 0)
                {
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
                    AudioSettingService.preferredVoices.Add(voice);
                }
            }
        }
    }
}
