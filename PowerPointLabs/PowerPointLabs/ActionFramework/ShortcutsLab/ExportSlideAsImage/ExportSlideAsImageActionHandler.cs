using System.Collections.Generic;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ShortcutsLab;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportActionRibbonId(ShortcutsLabText.ExportSlideAsImageTag)]
    class ExportSlideAsImageActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

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
            string savedFile = SaveFileDialogUtil.Save(
                ShortcutsLabConstants.ExportSlideSaveFileDialogTitle,
                ShortcutsLabConstants.ExportSlideSaveFileDialogFilter,
                ShortcutsLabConstants.ExportSlideSaveFileDialogExtension);
            if (savedFile != null)
            {           
                GraphicsUtil.ExportSlides(selectedSlides, savedFile);
            }

        }
    }
}
