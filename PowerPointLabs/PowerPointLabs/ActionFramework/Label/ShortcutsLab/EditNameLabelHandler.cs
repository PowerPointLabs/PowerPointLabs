using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        "EditNameMenuShape",
        "EditNameMenuLine",
        "EditNameMenuFreeform",
        "EditNameMenuPicture",
        "EditNameMenuGroup",
        "EditNameMenuChart",
        "EditNameMenuTable")]
    class EditNameLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.EditNameShapeLabel;
        }
    }
}
