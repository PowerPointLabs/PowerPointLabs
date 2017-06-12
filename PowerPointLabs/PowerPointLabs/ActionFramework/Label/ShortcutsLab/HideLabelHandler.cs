using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        "HideShapeMenuShape",
        "HideShapeMenuLine",
        "HideShapeMenuFreeform",
        "HideShapeMenuPicture",
        "HideShapeMenuGroup",
        "HideShapeMenuChart",
        "HideShapeMenuTable")]
    class HideShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.HideSelectedShapeLabel;
        }
    }
}
