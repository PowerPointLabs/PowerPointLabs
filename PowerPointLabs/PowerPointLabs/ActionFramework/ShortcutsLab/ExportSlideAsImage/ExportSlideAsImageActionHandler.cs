using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Models;
using PowerPointLabs.ShortcutsLab;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportActionRibbonId(ShortcutsLabText.ExportSlideAsImageTag)]
    class ExportSlideAsImageActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();


            // Get list of selected slides
            // Open save file dialog
            // try save
                // if only one, then save directly
                // if more than one, then save inside a folder with the specified name

            List<Slide> selectedSlides = new List<Slide>();

           

            if (this.GetCurrentSelection().Type == PpSelectionType.ppSelectionSlides)
            {
                foreach (Slide slide in this.GetCurrentSelection().SlideRange)
                {
                    selectedSlides.Add(slide);
                }
            }
            else
            {
                selectedSlides.Add(this.GetCurrentSlide().GetNativeSlide());
            }

            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog.Title = "Export Slide As Image";
            saveFileDialog.DefaultExt = "png";
            saveFileDialog.Filter = "Images|*.png;*.bmp;*.jpg";

            System.Windows.Forms.DialogResult result = saveFileDialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                GraphicsUtil.ExportSlides(selectedSlides, saveFileDialog.FileName);
            }

        }
    }
}
