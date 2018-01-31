using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ShortcutsLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.PasteLab
{
    [ExportActionRibbonId(ShortcutsLabText.FillSlideTag)]
    class FillSlideActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            Microsoft.Office.Interop.PowerPoint.Selection currentSelection = this.GetCurrentSelection();
            PowerPointLabs.Models.PowerPointSlide currentSlide = this.GetCurrentSlide();
            float slideHeight = this.GetCurrentPresentation().SlideHeight;
            float slideWidth = this.GetCurrentPresentation().SlideWidth;
            // Send over to Fill operation to fill up the slide
            FillSlide.Fill(currentSelection, currentSlide, slideWidth, slideHeight);
        }
    }
}
