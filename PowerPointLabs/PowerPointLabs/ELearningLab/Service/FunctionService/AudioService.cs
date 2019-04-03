using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using NAudio.Wave;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.ELearningLab.Utility;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.ELearningLab.Service
{
    public class AudioService
    {
#pragma warning disable 0618
        public static string TempFolderName
        {
            get
            {
                return string.Format(ELearningLabText.TempFolderNameFormat, 
                    Globals.ThisAddIn.GetTempFolderName());
            }
        }

        public static Effect CreateAppearEffectAudioAnimation(PowerPointSlide slide, string captionText, string voiceLabel,
           int clickNo, int tagNo, bool isSeperateClick)
        {
            Shape shape;
            try
            {
                shape = InsertAudioShapeToSlide(slide, captionText, tagNo, voiceLabel);
            }
            catch (Exception e)
            {
                Logger.Log(e.Message);
                return null;
            }
            Effect effect;
            if (shape == null)
            {
                return null;
            }
            if (isSeperateClick)
            {
                effect = AnimationUtility.AppendAnimationToSlide(slide, shape, MsoAnimEffect.msoAnimEffectMediaPlay, clickNo - 1);
            }
            else
            {
                effect = AnimationUtility.AppendAnimationToSlide(slide, shape, MsoAnimEffect.msoAnimEffectMediaPlay, clickNo);
            }
            // TODO: add stop playing on click here
            return effect;
        }

        public static IVoice GetVoiceFromString(string str)
        {
            if (Enum.IsDefined(typeof(AzureVoiceType), str))
            {
                return AzureVoiceList.voices.Where(x => x.Voice.ToString() == str).ElementAtOrDefault(0);
            }
            else if (Enum.IsDefined(typeof(WatsonVoiceType), str))
            {
                return WatsonRuntimeService.Voices.Where(x => x.Voice.ToString() == str).ElementAtOrDefault(0);
            }
            else
            {
                return ComputerVoiceRuntimeService.Voices.Where(x => x.Voice.ToString() == str).ElementAtOrDefault(0);
            }
        }

        public static VoiceType GetVoiceTypeFromString(string voiceName, string defaultPostfix)
        {
            if (IsDefaultVoiceType(defaultPostfix))
            {
                return VoiceType.DefaultVoice;
            }
            IVoice voice = GetVoiceFromString(voiceName);
            if (voice is AzureVoice)
            {
                return VoiceType.AzureVoice;
            }
            else if (voice is WatsonVoice)
            {
                return VoiceType.WatsonVoice;
            }
            else
            {
                return VoiceType.ComputerVoice;
            }
        }

        public static bool IsAzureVoiceSelectedForItem(ExplanationItem selfExplanationClickItem)
        {
            string voiceName = StringUtility.ExtractVoiceNameFromVoiceLabel(selfExplanationClickItem.VoiceLabel);
            IVoice voice = GetVoiceFromString(voiceName);
            return voice is AzureVoice;
        }

        public static bool IsWatsonVoiceSelectedForItem(ExplanationItem selfExplanationClickItem)
        {
            string voiceName = StringUtility.ExtractVoiceNameFromVoiceLabel(selfExplanationClickItem.VoiceLabel);
            IVoice voice = GetVoiceFromString(voiceName);
            return voice is WatsonVoice;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="fileName">Absolute path of the .wav file</param>
        /// <returns></returns>
        public static Shape InsertAudioFileOnSlide(PowerPointSlide slide, string fileName)
        {
            if (!File.Exists(fileName))
            {
                return null;
            }

            float slideWidth = PowerPointPresentation.Current.SlideWidth;

            Microsoft.Office.Interop.PowerPoint.Shapes shapes = slide.Shapes;

            try
            {
                Shape audioShape = slide.Shapes.AddMediaObject2(fileName, MsoTriState.msoFalse, MsoTriState.msoTrue, slideWidth + 20);
                slide.RemoveAnimationsForShape(audioShape);

                return audioShape;
            }
            catch 
            {
                Logger.Log("Audio not generated because text is not in English.");
                return null;
            }
        }

        public static TimeSpan ReadWavFileTimeSpan(string filepath)
        {
            WaveFileReader reader = new WaveFileReader(filepath);
            TimeSpan time = reader.TotalTime;
            return time;
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="captionText"></param>
        /// <param name="tagNo"></param>
        /// <param name="voiceLabel">Format "voiceName(_Default)"</param>
        /// <returns></returns>
        private static Shape InsertAudioShapeToSlide(PowerPointSlide slide, string captionText, int tagNo, string voiceLabel)
        {
            if (string.IsNullOrEmpty(captionText.Trim()))
            {
                return null;
            }
            string shapeName = string.Format(ELearningLabText.AudioCustomShapeNameFormat, tagNo, voiceLabel);
            if (!Directory.Exists(Path.Combine(Path.GetTempPath(), TempFolderName)))
            {
                Directory.CreateDirectory(Path.Combine(Path.GetTempPath(), TempFolderName));
            }
            string audioFilePath = Path.Combine(Path.GetTempPath(), TempFolderName,
                string.Format(ELearningLabText.AudioFileNameFormat, slide.ID, tagNo));
            string voiceName = StringUtility.ExtractVoiceNameFromVoiceLabel(voiceLabel);
            IVoice voice = GetVoiceFromString(voiceName);
            Shape shape = null;
            bool isSavedSuccessful = false;
            bool isSameCaptionText = IsSameCaptionText(slide, captionText, voiceLabel, tagNo);
            if (!isSameCaptionText)
            {
                isSavedSuccessful = SaveTextToWaveFile(captionText, audioFilePath, voice);
            }
            if (isSameCaptionText || isSavedSuccessful)
            {
                shape = InsertAudioFileOnSlide(slide, audioFilePath);
            }
            if (shape != null)
            {
                shape.Name = shapeName;
            }
            return shape;

        }

        private static bool SaveTextToWaveFile(string text, string filePath, IVoice voice)
        {
            if (voice is AzureVoice)
            {
                return AzureRuntimeService.SaveStringToWaveFileWithAzureVoice(text, filePath, voice as AzureVoice);            
            }
            else if (voice is ComputerVoice)
            {
                ComputerVoiceRuntimeService.SaveStringToWaveFile(text, filePath, voice as ComputerVoice);
                return true;
            }
            else if (voice is WatsonVoice)
            {
                WatsonRuntimeService.SaveStringToWaveFile(text, filePath, voice as WatsonVoice);
                return true;
            }
            return false;
        }

        private static bool IsDefaultVoiceType(string str)
        {
            return str.Equals(ELearningLabText.DefaultAudioIdentifier);
        }

        private static bool IsSameCaptionText(PowerPointSlide slide, string captionText, string voiceLabel, int tagNo)
        {
            string captionShapeName = string.Format(ELearningLabText.CaptionShapeNameFormat, tagNo);
            string audioShapeName = string.Format(ELearningLabText.AudioCustomShapeNameFormat, tagNo, voiceLabel);
            if (slide.HasShapeWithSameName(audioShapeName) && slide.HasShapeWithSameName(captionShapeName))
            {
                Shape shape = slide.GetShapeWithName(captionShapeName)[0];
                return captionText.Trim().Equals(shape.TextFrame.TextRange.Text.Trim());
            }
            return false;
        }
    }
}
