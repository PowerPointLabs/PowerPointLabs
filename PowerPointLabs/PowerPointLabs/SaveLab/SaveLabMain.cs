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

        public class PresentationDocument
        {
            public int PresentationDocID { get; set; }
            public string PresentationDocName { get; set; }
            public byte[] PresentationDocContent { get; set; }
        }
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

        public static void SaveFile(SlideRange selectedSlides)
        {
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

        }

        private 
    }
}
