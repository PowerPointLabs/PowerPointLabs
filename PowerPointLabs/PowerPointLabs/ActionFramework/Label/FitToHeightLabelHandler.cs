using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        "fitToHeightShape",
        "fitToHeightFreeform",
        "fitToHeightPicture",
        "fitToHeightChart",
        "fitToHeightTable")]
    class FitToHeightLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.FitToHeightShapeLabel;
        }
    }
}
