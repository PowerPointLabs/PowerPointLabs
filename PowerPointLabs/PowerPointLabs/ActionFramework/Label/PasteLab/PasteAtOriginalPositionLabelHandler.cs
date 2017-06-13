using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.PasteLab
{
    [ExportLabelRibbonId(
        "PasteAtOriginalPositionMenuShape", "PasteAtOriginalPositionMenuLine", "PasteAtOriginalPositionMenuFreeform",
        "PasteAtOriginalPositionMenuPicture", "PasteAtOriginalPositionMenuGroup", "PasteAtOriginalPositionMenuInk",
        "PasteAtOriginalPositionMenuVideo", "PasteAtOriginalPositionMenuTextEdit", "PasteAtOriginalPositionMenuChart",
        "PasteAtOriginalPositionMenuTable", "PasteAtOriginalPositionMenuTableWhole", "PasteAtOriginalPositionMenuFrame",
        "PasteAtOriginalPositionMenuSmartArtBackground", "PasteAtOriginalPositionMenuSmartArtEditSmartArt",
        "PasteAtOriginalPositionMenuSmartArtEditText")]
    class PasteAtOriginalPositionLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.PasteLabText.PasteAtOriginalPosition;
        }
    }
}