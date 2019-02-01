using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading;
using System.Windows.Forms;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.AudioMisc;
using PowerPointLabs.Models;
using PowerPointLabs.NarrationsLab.Data;
using PowerPointLabs.NarrationsLab.ViewModel;
using PowerPointLabs.SpeechEngine;
using PowerPointLabs.Views;

using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.NarrationsLab
{
    internal static class NotesToAudio
    {
#pragma warning disable 0618

        public const string SpeechShapePrefix = "PowerPointLabs Speech";
        public const string SpeechShapePrefixOld = "AudioGen Speech";

        public static bool IsRemoveAudioEnabled { get; set; } = true;
        public static bool IsAzureVoiceSelected { get; set; } = false;

        private static string TempFolderName
        {
            get
            {
                string tempName = Globals.ThisAddIn.GetActiveWindowTempName();
                return @"\PowerPointLabs Temp\" + tempName + @"\";
            }
        }

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

        public static List<string[]> EmbedSelectedSlideNotes()
        {
            ProcessingStatusForm progressBarForm = new ProcessingStatusForm();
            progressBarForm.Show();
            List<string[]> audioList = new List<string[]>();

            List<PowerPointSlide> slides = PowerPointCurrentPresentationInfo.SelectedSlides.ToList();
            if (NotesToAudio.IsAzureVoiceSelected && UserAccount.GetInstance().IsEmpty())
            {
                MessageBox.Show("Invalid user account. Please log in again.");
                throw new Exception("Invalid user account.");
            }
            int numberOfSlides = slides.Count;
            for (int currentSlideIndex = 0; currentSlideIndex < numberOfSlides; currentSlideIndex++)
            {
                int percentage = (int)Math.Round(((double)currentSlideIndex + 1) / numberOfSlides * 100);
                progressBarForm.UpdateProgress(percentage);
                progressBarForm.UpdateSlideNumber(currentSlideIndex, numberOfSlides);

                PowerPointSlide slide = slides[currentSlideIndex];
                audioList.Add(EmbedSlideNotes(slide));
            }
            progressBarForm.Close();

            return audioList;
        }

        public static List<string[]> EmbedAllSlideNotes()
        {
            ProcessingStatusForm progressBarForm = new ProcessingStatusForm();
            progressBarForm.Show();
            List<string[]> audioList = new List<string[]>();

            List<PowerPointSlide> slides = PowerPointPresentation.Current.Slides;

            int numberOfSlides = slides.Count;
            for (int currentSlideIndex = 0; currentSlideIndex < numberOfSlides; currentSlideIndex++)
            {
                int percentage = (int)Math.Round(((double)currentSlideIndex + 1) / numberOfSlides * 100);
                progressBarForm.UpdateProgress(percentage);
                progressBarForm.UpdateSlideNumber(currentSlideIndex, numberOfSlides);

                PowerPointSlide slide = slides[currentSlideIndex];
                audioList.Add(EmbedSlideNotes(slide));
            }
            progressBarForm.Close();

            return audioList;
        }

        public static bool OutputSlideNotesToFiles(PowerPointSlide slide, String folderPath)
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

        public static void SpeakSelectedText()
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

                SpeakText(completeTextBuilder.ToString());
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

        public static IEnumerable<String> GetVoices()
        {
            return TextToSpeech.GetVoices();
        }
        public static void SetDefaultVoice(string voiceName)
        {
            TextToSpeech.DefaultVoiceName = voiceName;
        }

        public static void SetDefaultVoice(string voiceName, AzureVoice humanVoice)
        {
            TextToSpeech.DefaultVoiceName = voiceName;
            TextToSpeech.humanVoice = humanVoice;
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
            DialogResult result = audioPicker.ShowDialog();

            if (result == DialogResult.OK)
            {
                string selectedFile = audioPicker.FileName;

                PowerPointSlide currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide;
                Shape newAudio = InsertAudioFileOnSlide(currentSlide, selectedFile);

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
            String folderPath = Path.GetTempPath() + TempFolderName;
            String fileNameSearchPattern = String.Format("Slide {0} Speech", slide.ID);

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
                    String fileName = audioFiles[i];
                    bool isOnClick = fileName.Contains("OnClick");

                    try
                    {
                        Shape audioShape = InsertAudioFileOnSlide(slide, fileName);
                        audioShape.Name = String.Format("PowerPointLabs Speech {0}", i);
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

        private static Shape InsertAudioFileOnSlide(PowerPointSlide slide, string fileName)
        {
            float slideWidth = PowerPointPresentation.Current.SlideWidth;

            Shape audioShape = slide.Shapes.AddMediaObject2(fileName, MsoTriState.msoFalse, MsoTriState.msoTrue, slideWidth + 20);
            slide.RemoveAnimationsForShape(audioShape);

            return audioShape;
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

        private static void SpeakText(string textToSpeak)
        {
            try
            {
                TextToSpeech.SpeakString(textToSpeak);
            }
            catch (InvalidOperationException)
            {
                ErrorParsingText();
            }
        }

        private static void ErrorParsingText()
        {
            MessageBox.Show(TextCollection.NarrationsLabText.RecorderErrorCannotParseText, 
                            TextCollection.NarrationsLabText.RecorderErrorCannotParseTextTitle,
                            MessageBoxButtons.OK, 
                            MessageBoxIcon.Error);
        }
    }
}
