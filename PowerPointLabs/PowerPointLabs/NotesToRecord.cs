using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using PowerPointLabs.AudioMisc;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs
{
    class NotesToRecord
    {
        public static void ReplaceArtificialWithVoice(Audio audio)
        {
            var currentSlide = PowerPointPresentation.CurrentSlide;
            var shapeName = audio.Name;
            var isOnClick = audio.SaveName.Contains("OnClick");

            // init click number
            var temp = audio.Name.Split(new[] {' '});
            var clickNumber = Int32.Parse(temp[2]);

            if (currentSlide != null)
            {
                // delete old shape
                currentSlide.DeleteShapesWithPrefix(shapeName);

                // embed new shape
                try
                {
                    var audioShape = InsertAudioFileOnSlide(currentSlide, audio.SaveName);
                    currentSlide.RemoveAnimationsForShape(audioShape);
                    audioShape.Name = shapeName;

                    if (isOnClick)
                    {
                        currentSlide.SetShapeAsClickTriggered(audioShape, clickNumber, MsoAnimEffect.msoAnimEffectMediaPlay);
                    }
                    else
                    {
                        currentSlide.SetAudioAsAutoplay(audioShape);
                    }
                }
                catch (COMException)
                {
                    // Adding the file failed for one reason or another - probably cancelled by the user.
                }
                MessageBox.Show("Record Replaced");
            }
            else
            {
                MessageBox.Show("Must select a slide");
            }
        }

        private static Shape InsertAudioFileOnSlide(PowerPointSlide slide, string fileName)
        {
            float slideWidth = PowerPointPresentation.SlideWidth;

            Shape audioShape = slide.Shapes.AddMediaObject2(fileName, MsoTriState.msoFalse, MsoTriState.msoTrue, slideWidth + 20);
            slide.RemoveAnimationsForShape(audioShape);

            return audioShape;
        }
    }
}
