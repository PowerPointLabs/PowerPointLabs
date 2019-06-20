using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Speech.Synthesis;
using System.Text;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.AudioMisc;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.Models;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

using OpenFileDialog = System.Windows.Forms.OpenFileDialog;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.ELearningLab.Service
{
    internal static class ComputerVoiceRuntimeService
    {
#pragma warning disable 0618

        public const string SpeechShapePrefix = "PowerPointLabs Speech";
        public const string SpeechShapePrefixOld = "AudioGen Speech";

        public static bool IsRemoveAudioEnabled { get; set; } = true;
        public static bool IsAzureVoiceSelected { get; set; } = false;

        public static ObservableCollection<ComputerVoice> Voices = GetVoices();


        #region Public Helper Functions

        public static void SpeakString(string textToSpeak, ComputerVoice voice)
        {
            if (string.IsNullOrWhiteSpace(textToSpeak))
            {
                return;
            }

            PromptBuilder builder = GetPromptForText(textToSpeak, voice);
            PromptToAudio.Speak(builder);
        }
        public static void SaveStringToWaveFile(string textToSave, string filePath, ComputerVoice voice)
        {
            PromptBuilder builder = GetPromptForText(textToSave, voice);
            PromptToAudio.SaveAsWav(builder, filePath);
        }


        #endregion

        #region Old NarrationsLab functions

        public static void PreviewAnimations()
        {
            try
            {
                Globals.ThisAddIn.Application.CommandBars.ExecuteMso("AnimationPreview");
            }
            catch (COMException)
            {
                // There wasn't anything to preview.
            }
        }

        public static string[] EmbedCurrentSlideNotes()
        {
            PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;

            if (currentSlide != null)
            {
                return EmbedSlideNotes(currentSlide);
            }

            return null;
        }

        public static void EmbedSelectedSlideNotes()
        {
            List<PowerPointSlide> slides = PowerPointCurrentPresentationInfo.SelectedSlides.ToList();

            if (AudioSettingService.selectedVoiceType == AudioGenerator.VoiceType.AzureVoice
                && AzureAccount.GetInstance().IsEmpty())
            {
                MessageBoxUtil.Show("Invalid user account. Please log in again.");
                throw new Exception("Invalid user account.");
            }

            int numberOfSlides = slides.Count;

            ProcessingStatusForm progressBarForm =
                new ProcessingStatusForm(numberOfSlides, BackgroundWorkerType.AudioGenerationService);
            progressBarForm.Show();
        }

        public static void EmbedSlideNotes(int i)
        {
            List<PowerPointSlide> slides = PowerPointCurrentPresentationInfo.SelectedSlides.ToList();
            PowerPointSlide slide = slides.ElementAt(i);
            EmbedSlideNotes(slide);
        }

        public static List<string[]> ExtractSlideNotes()
        {
            List<string[]> slideNotes = new List<string[]>();
            List<PowerPointSlide> slides = PowerPointCurrentPresentationInfo.SelectedSlides.ToList();
            string folderPath = Path.GetTempPath() + AudioService.TempFolderName;
            foreach (PowerPointSlide slide in slides)
            {
                string fileNameSearchPattern = String.Format("Slide {0} Speech", slide.ID);
                slideNotes.Add(GetAudioFilePaths(folderPath, fileNameSearchPattern));
            }
            return slideNotes;
        }

        public static bool OutputSlideNotesToFiles(PowerPointSlide slide, string folderPath)
        {
            try
            {
                String fileNameFormat = "Slide " + slide.ID + " Speech {0}";
                TextToSpeech.SaveStringToWaveFiles(slide.NotesPageText, folderPath, fileNameFormat);
                return true;
            }
            catch (InvalidOperationException)
            {
                ErrorParsingText();
            }
            return false;
        }

        public static void SpeakSelectedText(ComputerVoice voice)
        {
            try
            {
                string selected = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange.Text.Trim();
                List<string> splitScript = (new TaggedText(selected)).SplitByClicks();

                StringBuilder completeTextBuilder = new StringBuilder();
                Regex reg = new Regex("\\.+\\s*");

                foreach (string text in splitScript)
                {
                    completeTextBuilder.Append(reg.Replace(text, string.Empty));
                    completeTextBuilder.Append(". ");
                }

                SpeakText(completeTextBuilder.ToString(), voice);
            }
            catch (COMException)
            {
                // Nothing was selected.
            }
        }

        public static void RemoveAudioFromSelectedSlides()
        {
            foreach (PowerPointSlide s in PowerPointCurrentPresentationInfo.SelectedSlides)
            {
                s.DeleteShapesWithPrefixTimelineInvariant(SpeechShapePrefix);
                s.DeleteShapesWithPrefixTimelineInvariant(SpeechShapePrefixOld);
            }
        }

        public static void ReplaceSelectedAudio()
        {
            Microsoft.Office.Interop.PowerPoint.ShapeRange selectedShape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
            if (selectedShape.Count != 1 || selectedShape.MediaType != PpMediaType.ppMediaTypeSound)
            {
                return;
            }

            OpenFileDialog audioPicker = new OpenFileDialog
            {
                Filter = "Audio files (*.wav, *.mp3, *.wma)|*.wav;*.mp3;*.wma"
            };
            System.Windows.Forms.DialogResult result = audioPicker.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                string selectedFile = audioPicker.FileName;

                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
                Shape newAudio = AudioService.InsertAudioFileOnSlide(currentSlide, selectedFile);

                currentSlide.TransferAnimation(selectedShape[1], newAudio);

                selectedShape.Delete();
            }
        }


        /// <summary>
        /// This function will embed the auto generated speech to the current slide.
        /// File names of generated audios will be returned.
        /// </summary>
        /// <param name="slide">Current slide reference.</param>
        /// <returns>An array of auto generated audios' name.</returns>
        private static string[] EmbedSlideNotes(PowerPointSlide slide)
        {
            string folderPath = Path.GetTempPath() + AudioService.TempFolderName;

            string fileNameSearchPattern = string.Format("Slide {0} Speech", slide.ID);

            Directory.CreateDirectory(folderPath);

            // TODO:
            // obviously deleting all audios in current slide may not a good idea, some lines of script
            // may still be the same. Check the line first before deleting, if the line has not been
            // changed, leave the audio.

            // to avoid duplicate records, delete all old audios in the current slide
            string[] audiosInCurrentSlide = Directory.GetFiles(folderPath);
            foreach (string audio in audiosInCurrentSlide)
            {
                if (audio.Contains(fileNameSearchPattern))
                {
                    try
                    {
                        File.Delete(audio);
                    }
                    catch (Exception e)
                    {
                        Logger.LogException(e, "Failed to delete audio, it may be still playing. " + e.Message);
                    }
                }
            }

            bool isSaveSuccessful = OutputSlideNotesToFiles(slide, folderPath);
            string[] audioFiles = null;

            if (isSaveSuccessful)
            {
                slide.DeleteShapesWithPrefix(SpeechShapePrefix);

                audioFiles = GetAudioFilePaths(folderPath, fileNameSearchPattern);

                for (int i = 0; i < audioFiles.Length; i++)
                {
                    string fileName = audioFiles[i];
                    bool isOnClick = fileName.Contains("OnClick");

                    try
                    {
                        Shape audioShape = AudioService.InsertAudioFileOnSlide(slide, fileName);
                        audioShape.Name = string.Format("PowerPointLabs Speech {0}", i);
                        slide.RemoveAnimationsForShape(audioShape);

                        if (isOnClick)
                        {
                            slide.SetShapeAsClickTriggered(audioShape, i, MsoAnimEffect.msoAnimEffectMediaPlay);
                        }
                        else
                        {
                            slide.SetAudioAsAutoplay(audioShape);
                        }
                    }
                    catch (COMException)
                    {
                        // Adding the file failed for one reason or another - probably cancelled by the user.
                    }
                }
            }

            return audioFiles;
        }

        private static string[] GetAudioFilePaths(string folderPath, string fileNameSearchPattern)
        {
            IEnumerable<string> filePaths = Directory.EnumerateFiles(folderPath, "*." + Audio.RecordedFormatExtension);
            Utils.Comparers.AtomicNumberStringCompare comparer = new Utils.Comparers.AtomicNumberStringCompare();
            string[] audioFiles =
                filePaths.Where(path => path.Contains(fileNameSearchPattern)).OrderBy(x => new FileInfo(x).Name,
                                                                                      comparer).ToArray();

            return audioFiles;
        }

        private static void SpeakText(string textToSpeak, ComputerVoice voice)
        {
            try
            {
                SpeakString(textToSpeak, voice);
            }
            catch (InvalidOperationException)
            {
                ErrorParsingText();
            }
        }

        private static void ErrorParsingText()
        {
            MessageBoxUtil.Show(TextCollection.NarrationsLabText.RecorderErrorCannotParseText,
                            TextCollection.NarrationsLabText.RecorderErrorCannotParseTextTitle,
                            MessageBoxButtons.OK,
                            MessageBoxIcon.Error);
        }

        #endregion

        #region Private Helper Functions
        private static PromptBuilder GetPromptForText(string textToConvert, ComputerVoice voice)
        {
            TaggedText taggedText = new TaggedText(textToConvert);
            PromptBuilder builder = taggedText.ToPromptBuilder(voice.ToString());
            return builder;
        }

        private static ObservableCollection<ComputerVoice> GetVoices()
        {
            ObservableCollection<ComputerVoice> voices = new ObservableCollection<ComputerVoice>();
            List<string> installedVoices = TextToSpeech.GetVoices().ToList();
            foreach (string voiceName in installedVoices)
            {
                voices.Add(new ComputerVoice(voiceName));
            }
            return voices;
        }
        #endregion
    }
}
