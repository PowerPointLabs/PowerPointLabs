using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId(
        "fitToWidthShape",
        "fitToWidthFreeform",
        "fitToWidthPicture",
        "fitToWidthChart",
        "fitToWidthTable")]
    class FitToWidthActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();
            var selectedShape = this.GetCurrentSelection().ShapeRange[1];
            var pres = this.GetCurrentPresentation();
            FitToSlide.FitToWidth(selectedShape, pres.SlideWidth, pres.SlideHeight);
        }
    }
}
