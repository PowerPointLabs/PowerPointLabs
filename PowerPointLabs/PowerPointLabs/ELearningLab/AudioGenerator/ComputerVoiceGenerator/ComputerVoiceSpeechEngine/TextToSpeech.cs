using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Speech.Synthesis;
using System.Text;
using System.Threading;
using System.Windows;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.Models;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    static class TextToSpeech
    {
        public static IEnumerable<string> GetVoices()
        {
            using (SpeechSynthesizer synthesizer = new SpeechSynthesizer())
            {
                System.Collections.ObjectModel.ReadOnlyCollection<InstalledVoice> installedVoices = synthesizer.GetInstalledVoices();
                IEnumerable<InstalledVoice> voices = installedVoices.Where(voice => voice.Enabled);
                return voices.Select(voice => voice.VoiceInfo.Name);
            }
        }

        public static void SaveStringToWaveFiles(string notesText, string folderPath, string fileNameFormat)
        {
            TaggedText taggedNotes = new TaggedText(notesText);
            List<string> stringsToSave = taggedNotes.SplitByClicks();
            //MD5 md5 = MD5.Create();

            for (int i = 0; i < stringsToSave.Count; i++)
            {
                string textToSave = stringsToSave[i];
                string baseFileName = string.Format(fileNameFormat, i + 1);

                // The first item will autoplay; everything else is triggered by a click.
                string fileName = i > 0 ? baseFileName + " (OnClick)" : baseFileName;
                string filePath = folderPath + "\\" + fileName + ".wav";

                switch (AudioSettingService.selectedVoiceType)
                {
                    case VoiceType.ComputerVoice:
                        ComputerVoiceRuntimeService.SaveStringToWaveFile(textToSave, filePath,
                            AudioSettingService.selectedVoice as ComputerVoice);
                        break;
                    case VoiceType.AzureVoice:
                        AzureRuntimeService.SaveStringToWaveFileWithAzureVoice(textToSave, filePath, 
                            AudioSettingService.selectedVoice as AzureVoice);
                        break;
                    default:
                        break;
                }
            }
        }
    }
}