using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        "FitToHeightMenuShape",
        "FitToHeightMenuFreeform",
        "FitToHeightMenuPicture",
        "FitToHeightMenuGroup",
        "FitToHeightMenuChart",
        "FitToHeightMenuTable")]
    class FitToHeightLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.FitToHeightShapeLabel;
        }
    }
}
