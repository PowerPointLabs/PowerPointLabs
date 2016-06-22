﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Security.Cryptography;
using System.Speech.Synthesis;
using System.Text;
using PowerPointLabs.Models;
using System.Text.RegularExpressions;

namespace PowerPointLabs.SpeechEngine
{
    static class TextToSpeech
    {
        public static String DefaultVoiceName;

        public static void SpeakString(String textToSpeak)
        {
            if (String.IsNullOrWhiteSpace(textToSpeak))
            {
                return;
            }

            var newTextToSpeak = ReadSpelledOutWord(textToSpeak);

            var builder = GetPromptForText(newTextToSpeak);
            PromptToAudio.Speak(builder);
        }

        public static string ReadSpelledOutWord(String textToSpeak)
        {
            string space = " ";
            var textToSpeakList = textToSpeak.Split(space.ToCharArray()[0]);
            string newTextToSpeak = "";

            for (int i = 0; i < textToSpeakList.Length; i++)
            {
                var thisWord = textToSpeakList[i];
                var charList = thisWord.ToArray();
                bool allUpper = true;
                if (thisWord.Length >= 2)
                {

                    for (int j = 0; j < charList.Length; j++)
                    {
                        var thisChar = charList[j];

                        if (!Regex.IsMatch(thisChar.ToString(), "[A-Z]"))
                        {
                            if (!Char.IsPunctuation(thisChar))
                            {
                                allUpper = false;
                            }
                        }
                    }

                    if (allUpper)
                    {
                        thisWord = "";
                        for (int j = 0; j < charList.Length; j++)
                        {
                            var thisChar = charList[j];
                            if (thisChar.Equals("."))
                            {
                                thisWord = thisWord + ";" + " ";
                            }
                            else
                            {
                                thisWord = thisWord + " " + thisChar.ToString();
                            }
                        }
                    }
                }
                newTextToSpeak = newTextToSpeak + " " + thisWord;
            }

            return newTextToSpeak;
        }

        private static PromptBuilder GetPromptForText(string textToConvert)
        {
            var taggedText = new TaggedText(textToConvert);
            PromptBuilder builder = taggedText.ToPromptBuilder(DefaultVoiceName);
            return builder;
        }

        public static void SaveStringToWaveFiles(string notesText, string folderPath, string fileNameFormat)
        {
            var taggedNotes = new TaggedText(notesText);
            List<String> stringsToSave = taggedNotes.SplitByClicks();
            //MD5 md5 = MD5.Create();

            for (int i = 0; i < stringsToSave.Count; i++)
            {
                String textToSave = stringsToSave[i];
                String baseFileName = String.Format(fileNameFormat, i + 1);
                
                // The first item will autoplay; everything else is triggered by a click.
                String fileName = i > 0 ? baseFileName + " (OnClick)" : baseFileName;

                String filePath = folderPath + "\\" + fileName + ".wav";

                SaveStringToWaveFile(textToSave, filePath);
            }
        }

        public static void SaveStringToWaveFile(String textToSave, String filePath)
        {
            var newTextToSave = ReadSpelledOutWord(textToSave);
            var builder = GetPromptForText(newTextToSave);
            PromptToAudio.SaveAsWav(builder, filePath);
        }

        public static IEnumerable<string> GetVoices()
        {
            using (var synthesizer = new SpeechSynthesizer())
            {
                var installedVoices = synthesizer.GetInstalledVoices();
                var voices = installedVoices.Where(voice => voice.Enabled);
                return voices.Select(voice => voice.VoiceInfo.Name);
            }
        }
    }
}