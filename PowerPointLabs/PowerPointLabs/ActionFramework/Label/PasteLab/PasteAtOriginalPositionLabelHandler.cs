using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.PasteLab
{
    [ExportLabelRibbonId(
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuShape,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuLine,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuFreeform,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuPicture,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuGroup,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuInk,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuVideo,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuTextEdit,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuChart,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuTable,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuTableCell,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuSlide,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuSmartArt,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuEditSmartArt,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.MenuEditSmartArtText,
        TextCollection.PasteAtOriginalPositionMenuId + TextCollection.RibbonButton)]
    class PasteAtOriginalPositionLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.PasteLabText.PasteAtOriginalPosition;
        }
    }
}