using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.PasteLab
{
    [ExportLabelRibbonId(
        "PasteAtOriginalPositionMenuFrame",
        "PasteAtOriginalPositionMenuShape",
        "PasteAtOriginalPositionMenuLine",
        "PasteAtOriginalPositionMenuFreeform",
        "PasteAtOriginalPositionMenuPicture",
        "PasteAtOriginalPositionMenuGroup")]
    class PasteAtOriginalPositionLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.PasteLabText.PasteAtOriginalPosition;
        }
    }
}