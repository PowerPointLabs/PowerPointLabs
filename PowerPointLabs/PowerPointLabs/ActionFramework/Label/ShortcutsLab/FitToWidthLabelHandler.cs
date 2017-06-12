using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        "FitToWidthMenuShape",
        "FitToWidthMenuFreeform",
        "FitToWidthMenuPicture",
        "FitToWidthMenuGroup",
        "FitToWidthMenuChart",
        "FitToWidthMenuTable")]
    class FitToWidthLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.FitToWidthShapeLabel;
        }
    }
}
