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
        /*
        public PresentationDocument FileToByteArray(string fileName)
        {
            byte[] fileContent = null;
            FileStream fs = new FileStream(fileName, FileMode.Create, FileAccess.Read);
            System.IO.BinaryReader binaryReader = new System.IO.BinaryReader(fs);
            long byteLength = new System.IO.FileInfo(fileName).Length;
            fileContent = binaryReader.ReadBytes((Int32)byteLength);
            fs.Close();
            fs.Dispose();
            binaryReader.Close();
            Document = new Document();
            Document.DocName = fileName;
            Document.DocContent = fileContent;
            return Document;
        }
        */

        public static void SaveFile(Models.PowerPointPresentation currentPresentation, Selection currentSelection)
        {

            SaveFileDialog saveFileDialog = new SaveFileDialog();
            //Presentation newPresentation = new Presentation();
            Models.PowerPointPresentation newPresentation;
            //System.Diagnostics.Debug.WriteLine(newPresentation.Presentation == null);
            //newPresentation.AddSlide(PpSlideLayout.ppLayoutBlank, "TEMP_SLIDE");
            //System.Diagnostics.Debug.WriteLine(newPresentation.Presentation == null);
            //newPresentation.Presentation.Slides.Add(1, PpSlideLayout.ppLayoutBlank);

            List<Models.PowerPointSlide> selectedSlides = currentPresentation.SelectedSlides;

            saveFileDialog.InitialDirectory = SaveLabSettings.SaveFolderPath;
            saveFileDialog.Filter = "PowerPoint Presentations|*.ppt";
            saveFileDialog.Title = "Save Selected Slides";
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.OverwritePrompt = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                currentPresentation.Presentation.SaveCopyAs(saveFileDialog.FileName);
                Presentations newPres = new Microsoft.Office.Interop.PowerPoint.Application().Presentations;
                Presentation tempPresentation = newPres.Open(saveFileDialog.FileName, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
                newPresentation = new Models.PowerPointPresentation(tempPresentation);
                System.Diagnostics.Debug.WriteLine("No of slides selected: " + selectedSlides.Count);

                foreach (Models.PowerPointSlide slide in selectedSlides)
                {
                    newPresentation.AddSlide(PpSlideLayout.ppLayoutMixed, slide.Name);
                }
                System.Diagnostics.Debug.WriteLine("No of slides saved: " + newPresentation.SlideCount);
                //newPresentation.RemoveSlide("TEMP_SLIDE", true);
                //System.Diagnostics.Debug.WriteLine("No of slides saved after removal: " + newPresentation.SlideCount);
                newPresentation.Save();
            }
            /*
            FileStream fileStream;
            SaveFileDialog saveFileDialog = new SaveFileDialog();

            saveFileDialog.InitialDirectory = SaveLabSettings.SaveFolderPath;
            saveFileDialog.Filter = "PowerPoint Presentations|*.ppt";
            saveFileDialog.Title = "Save Selected Slides";
            saveFileDialog.RestoreDirectory = true;
            saveFileDialog.OverwritePrompt = true;

            if (saveFileDialog.ShowDialog() == DialogResult.OK)
            {
                
                if ((fileStream = (FileStream)saveFileDialog.OpenFile()) != null)
                {
                    // Code to write the stream goes here.
                    GetBytes(selectedSlides)
                    fileStream.Close();
                }
                
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
            */
        }
    }
}
