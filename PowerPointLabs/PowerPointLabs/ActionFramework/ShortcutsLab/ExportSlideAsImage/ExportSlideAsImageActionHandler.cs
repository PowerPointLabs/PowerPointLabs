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

            Slide selectedSlide;

            if (this.GetCurrentSelection().Type == PpSelectionType.ppSelectionSlides)
            {
                selectedSlide = this.GetCurrentSelection().SlideRange[1];
            }
            else
            {
                selectedSlide = (Slide)this.GetCurrentSlide();
            }

            System.Windows.Forms.SaveFileDialog saveFileDialog = new System.Windows.Forms.SaveFileDialog();
            saveFileDialog.DefaultExt = "png";
            saveFileDialog.Filter = "Images|*.png;*.bmp;*.jpg";
            saveFileDialog.Title = "Export Slide As Image";

            System.Windows.Forms.DialogResult result = saveFileDialog.ShowDialog();

            if (result == System.Windows.Forms.DialogResult.OK)
            {
                GraphicsUtil.ExportSlide(selectedSlide, saveFileDialog.FileName);
            }

           



        }
    }
}
