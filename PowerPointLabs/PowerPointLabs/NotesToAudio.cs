using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using AudioGen.Models;
using AudioGen.SpeechEngine;
using AudioGen.Views;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace AudioGen
{
    class NotesToAudio
    {
        private const string TempFolderName = "\\AudioGen Temp\\";
        private const string SpeechShapePrefix = "AudioGen Speech";

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

        public static void EmbedCurrentSlideNotes()
        {
            var currentSlide = PowerPointPresentation.CurrentSlide;
            if (currentSlide != null)
            {
                EmbedSlideNotes(currentSlide);
            }
        }

        public static void EmbedAllSlideNotes()
        {
            ProcessingStatusForm progressBarForm = new ProcessingStatusForm();
            progressBarForm.Show();

            var slides = PowerPointPresentation.Slides.ToList();

            int numberOfSlides = slides.Count;
            for (int currentSlideIndex = 0; currentSlideIndex < numberOfSlides; currentSlideIndex++)
            {
                int percentage = (int)Math.Round(((double)currentSlideIndex) / numberOfSlides * 100);
                progressBarForm.UpdateProgress(percentage);
                progressBarForm.UpdateSlideNumber(currentSlideIndex, numberOfSlides);

                var slide = slides[currentSlideIndex];
                EmbedSlideNotes(slide);
            }
            progressBarForm.Close();
        }

        private static void EmbedSlideNotes(PowerPointSlide slide)
        {
            String folderPath = Path.GetTempPath() + TempFolderName;
            Directory.CreateDirectory(folderPath);

            bool isSaveSuccessful = OutputSlideNotesToFiles(slide, folderPath);
            if (isSaveSuccessful)
            {
                slide.DeleteShapesWithPrefix(SpeechShapePrefix);

                String fileNameSearchPattern = String.Format("Slide {0} Speech", slide.Index);
                var audioFiles = GetAudioFilePaths(folderPath, fileNameSearchPattern);

                int clickCount = 0;
                for (int i = 0; i < audioFiles.Length; i++)
                {
                    String fileName = audioFiles[i];
                    bool isOnClick = fileName.Contains("OnClick");

                    try
                    {
                        Shape audioShape = InsertAudioFileOnSlide(slide, fileName);
                        audioShape.Name = String.Format("AudioGen Speech {0}", i);
                        slide.RemoveAnimationsForShape(audioShape);

                        if (isOnClick)
                        {
                            clickCount++;
                            slide.SetAudioAsClickTriggered(audioShape, clickCount);
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

            Directory.Delete(folderPath, true);
        }

        private static Shape InsertAudioFileOnSlide(PowerPointSlide slide, string fileName)
        {
            float slideWidth = PowerPointPresentation.SlideWidth;

            Shape audioShape = slide.Shapes.AddMediaObject2(fileName, MsoTriState.msoFalse, MsoTriState.msoTrue, slideWidth + 20);
            slide.RemoveAnimationsForShape(audioShape);

            return audioShape;
        }

        private static string[] GetAudioFilePaths(string folderPath, string fileNameSearchPattern)
        {
            var filePaths = Directory.EnumerateFiles(folderPath, "*.wav");
            var audioFiles = filePaths.Where(path => path.Contains(fileNameSearchPattern)).ToArray();
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
            MessageBox.Show("Have you added the correct closing tags? \n(Speed and Gender text ranges can't overlap.)", "Couldn't Parse Text",
                MessageBoxButtons.OK, MessageBoxIcon.Error);
        }

        public static bool OutputSlideNotesToFiles(PowerPointSlide slide, String folderPath)
        {
            try
            {
                String fileNameFormat = "Slide " + slide.Index + " Speech {0}";
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
                String selected = Globals.ThisAddIn.Application.ActiveWindow.Selection.TextRange.Text;
                SpeakText(selected);
            }
            catch (COMException)
            {
                // Nothing was selected.
            }
        }

        public static void RemoveAudioFromAllSlides()
        {
            var slides = PowerPointPresentation.Slides;
            foreach (PowerPointSlide s in slides)
            {
                s.DeleteShapesWithPrefix(SpeechShapePrefix);
            }
        }

        public static void RemoveAudioFromCurrentSlide()
        {
            var currentSlide = PowerPointPresentation.CurrentSlide;
            if (currentSlide == null)
            {
                return;
            }
            currentSlide.DeleteShapesWithPrefix(SpeechShapePrefix);
        }

        public static IEnumerable<String> GetVoices()
        {
            return TextToSpeech.GetVoices();
        }
        public static void SetDefaultVoice(string voiceName)
        {
            TextToSpeech.DefaultVoiceName = voiceName;
        }

        public static void ReplaceSelectedAudio()
        {
            var selectedShape = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
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
                var selectedFile = audioPicker.FileName;

                PowerPointSlide currentSlide = PowerPointPresentation.CurrentSlide;
                Shape newAudio = InsertAudioFileOnSlide(currentSlide, selectedFile);

                currentSlide.TransferAnimation(selectedShape[1], newAudio);
                
                selectedShape.Delete();
            }
        }
    }
}
