using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Microsoft.Office.Core;
using PowerPointLabs.Models;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs
{
    class NotesToRecord
    {
        public static void EmbedRecordToSlide(string recName)
        {
            var currentSlide = PowerPointPresentation.CurrentSlide;

            if (currentSlide != null)
            {
                currentSlide.DeleteShapesWithPrefix(recName);

                try
                {
                    Shape audioShape = InsertAudioFileOnSlide(currentSlide, recName);
                    currentSlide.RemoveAnimationsForShape(audioShape);
                    audioShape.Name = recName;
                    currentSlide.SetAudioAsAutoplay(audioShape);
                }
                catch (COMException)
                {
                    // Adding the file failed for one reason or another - probably cancelled by the user.
                }
                MessageBox.Show("Record Embeded");
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
