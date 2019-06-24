using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Utils.Windows;
using PowerPointLabs.Views;

namespace PowerPointLabs.SaveLab
{
    static class SaveLabMain
    {
        public static void SaveFile(Models.PowerPointPresentation currentPresentation)
        {
            // Opens up a new Save File Dialog
            string savedFile = SaveFileDialogUtil.Save("Save Selected Slides",
                "PowerPoint Presentations|*.pptx",
                initialDirectory: SaveLabSettings.GetSaveFolderPath());
            WPFSaveFileDialog saveFileDialog = new WPFSaveFileDialog();

            Models.PowerPointPresentation newPresentation;
            List<Models.PowerPointSlide> selectedSlides = currentPresentation.SelectedSlides;

            if (savedFile != null)
            {
                // Copy the Current Presentation under a new name
                currentPresentation.Presentation.SaveCopyAs(saveFileDialog.FileName, PpSaveAsFileType.ppSaveAsDefault);

                try
                {
                    // Re-open the save copy in the same directory in the background
                    Presentations newPres = new Microsoft.Office.Interop.PowerPoint.Application().Presentations;
                    Presentation tempPresentation = newPres.Open(saveFileDialog.FileName, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse);
                    newPresentation = new Models.PowerPointPresentation(tempPresentation);

                    // Hashset to store the unique IDs of selected slides
                    HashSet<int> idHash = new HashSet<int>();
                    foreach (Models.PowerPointSlide selectedSlide in selectedSlides)
                    {
                        idHash.Add(selectedSlide.ID);
                    }

                    // Check each slide in new presentation and remove un-selected slides using unique slide ID
                    for (int i = newPresentation.SlideCount - 1; i >= 0; i--)
                    {
                        if (!idHash.Contains(newPresentation.Slides[i].ID))
                        {
                            newPresentation.RemoveSlide(i);
                        }
                    }
                    
                    // Check for and remove empty sections in new presentation
                    if (newPresentation.HasEmptySection)
                    {
                        newPresentation.RemoveEmptySections();
                    }

                    // Save and then close the presentation
                    newPresentation.Save();
                    newPresentation.Close();
                }
                catch (System.Runtime.InteropServices.COMException)
                {
                    // do nothing as file is successfully copied
                }
            }
        }
    }
}
