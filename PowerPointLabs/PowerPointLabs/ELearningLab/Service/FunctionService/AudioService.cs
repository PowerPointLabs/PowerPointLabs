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
                string tempName = Globals.ThisAddIn.GetActiveWindowTempName();
                return @"\PowerPointLabs Temp\" + tempName + @"\";
            }
        }
        public static Effect CreateAppearEffectAudioAnimation(PowerPointSlide slide, string captionText, string voiceLabel,
           int clickNo, int tagNo, bool isSeperateClick)
        {
            Shape shape = InsertAudioShapeToSlide(slide, captionText, tagNo, voiceLabel);
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
            else
            {
                return VoiceType.ComputerVoice;
            }
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="fileName">Absolute path of the .wav file</param>
        /// <returns></returns>
        public static Shape InsertAudioFileOnSlide(PowerPointSlide slide, string fileName)
        {
            float slideWidth = PowerPointPresentation.Current.SlideWidth;

            Shape audioShape = slide.Shapes.AddMediaObject2(fileName, MsoTriState.msoFalse, MsoTriState.msoTrue, slideWidth + 20);
            slide.RemoveAnimationsForShape(audioShape);

            return audioShape;
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
            if (!IsSameCaptionText(slide, captionText, voiceLabel, tagNo))
            {
                SaveTextToWaveFile(captionText, audioFilePath, voice);
            }
            shape = InsertAudioFileOnSlide(slide, audioFilePath);
            shape.Name = shapeName;
           // ReadWavFileTimeSpan(audioFilePath);
            return shape;

        }

        private static bool SaveTextToWaveFile(string text, string filePath, IVoice voice)
        {
            if (voice is AzureVoice)
            {
                AzureRuntimeService.SaveStringToWaveFileWithAzureVoice(text, filePath, voice as AzureVoice);
                return true;
            }
            else if (voice is ComputerVoice)
            {
                ComputerVoiceRuntimeService.SaveStringToWaveFile(text, filePath, voice as ComputerVoice);
                return true;
            }
            return false;
        }

        private static bool IsDefaultVoiceType(string str)
        {
            return str.Equals(ELearningLabText.DefaultAudioIdentifier);
        }

        private static void ReadWavFileTimeSpan(string filepath)
        {
            WaveFileReader reader = new WaveFileReader(filepath);
            TimeSpan time = reader.TotalTime;
            Logger.Log("time is " + time.ToString());
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
