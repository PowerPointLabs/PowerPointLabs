using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.PasteLab
{
    [ExportLabelRibbonId(
        "PasteAtCursorPositionMenuFrame",
        "PasteAtCursorPositionMenuShape",
        "PasteAtCursorPositionMenuLine",
        "PasteAtCursorPositionMenuFreeform",
        "PasteAtCursorPositionMenuPicture",
        "PasteAtCursorPositionMenuGroup",
        "PasteAtCursorPositionMenuInk",
        "PasteAtCursorPositionMenuVideo",
        "PasteAtCursorPositionMenuChart",
        "PasteAtCursorPositionMenuTable",
        "PasteAtCursorPositionMenuTableWhole",
        "PasteAtCursorPositionMenuSmartArtBackground",
        "PasteAtCursorPositionMenuSmartArtEditSmartArt",
        "PasteAtCursorPositionMenuSmartArtEditText")]
    class PasteAtCursorPositionLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.PasteLabText.PasteAtCursorPosition;
        }
    }
}