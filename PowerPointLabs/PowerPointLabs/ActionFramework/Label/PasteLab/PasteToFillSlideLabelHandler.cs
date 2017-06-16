using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.PasteLab
{
    [ExportLabelRibbonId(
        TextCollection.PasteToFillSlideId + TextCollection.MenuShape,
        TextCollection.PasteToFillSlideId + TextCollection.MenuLine,
        TextCollection.PasteToFillSlideId + TextCollection.MenuFreeform,
        TextCollection.PasteToFillSlideId + TextCollection.MenuPicture,
        TextCollection.PasteToFillSlideId + TextCollection.MenuGroup,
        TextCollection.PasteToFillSlideId + TextCollection.MenuInk,
        TextCollection.PasteToFillSlideId + TextCollection.MenuVideo,
        TextCollection.PasteToFillSlideId + TextCollection.MenuTextEdit,
        TextCollection.PasteToFillSlideId + TextCollection.MenuChart,
        TextCollection.PasteToFillSlideId + TextCollection.MenuTable,
        TextCollection.PasteToFillSlideId + TextCollection.MenuTableCell,
        TextCollection.PasteToFillSlideId + TextCollection.MenuSlide,
        TextCollection.PasteToFillSlideId + TextCollection.MenuSmartArt,
        TextCollection.PasteToFillSlideId + TextCollection.MenuEditSmartArt,
        TextCollection.PasteToFillSlideId + TextCollection.MenuEditSmartArtText,
        TextCollection.PasteToFillSlideId + TextCollection.RibbonButton)]
    class PasteToFillSlideLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.PasteLabText.PasteToFillSlide;
        }
    }
}