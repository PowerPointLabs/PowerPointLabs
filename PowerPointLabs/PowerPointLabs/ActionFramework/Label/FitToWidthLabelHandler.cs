using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        "fitToWidthShape",
        "fitToWidthFreeform",
        "fitToWidthPicture",
        "fitToWidthChart",
        "fitToWidthTable")]
    class FitToWidthLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId, string ribbonTag)
        {
            return TextCollection.FitToWidthShapeLabel;
        }
    }
}
