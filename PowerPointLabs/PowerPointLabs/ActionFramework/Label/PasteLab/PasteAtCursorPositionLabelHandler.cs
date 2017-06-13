using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.PasteLab
{
    [ExportLabelRibbonId(
        "PasteAtCursorPositionMenuShape", "PasteAtCursorPositionMenuLine", "PasteAtCursorPositionMenuFreeform",
        "PasteAtCursorPositionMenuPicture", "PasteAtCursorPositionMenuGroup", "PasteAtCursorPositionMenuInk",
        "PasteAtCursorPositionMenuVideo", "PasteAtCursorPositionMenuTextEdit", "PasteAtCursorPositionMenuChart",
        "PasteAtCursorPositionMenuTable", "PasteAtCursorPositionMenuTableWhole", "PasteAtCursorPositionMenuFrame",
        "PasteAtCursorPositionMenuSmartArtBackground", "PasteAtCursorPositionMenuSmartArtEditSmartArt",
        "PasteAtCursorPositionMenuSmartArtEditText")]
    class PasteAtCursorPositionLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.PasteLabText.PasteAtCursorPosition;
        }
    }
}