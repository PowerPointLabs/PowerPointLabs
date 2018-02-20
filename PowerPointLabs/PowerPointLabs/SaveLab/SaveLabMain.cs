using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.SaveLab
{
    class SaveLabMain
    {
        public static void SaveFile(SlideRange selectedSlides)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.InitialDirectory = SaveLabSettings.SaveFolderPath;
            saveFileDialog.Filter = "PowerPoint Presentations|*.ppt";
            saveFileDialog.Title = "Save Selected Slides";
            saveFileDialog.RestoreDirectory = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                // Creates a new presentation and removes the first slide
                Presentation newPresentation = new Presentation();
                newPresentation.Slides.FindBySlideID(1).Delete();
                // Copies over the selected slides
                for (int i = 0; i < selectedSlides.Count; i++)
                {
                    newPresentation.Slides.AddSlide(i + 1, selectedSlides[1].CustomLayout);
                }
                //Save the new presentation as a new name
                newPresentation.SaveAs(saveFileDialog.FileName);
                System.Diagnostics.Process.Start(saveFileDialog.FileName + ".pptx");
            }

        }
    }
}
