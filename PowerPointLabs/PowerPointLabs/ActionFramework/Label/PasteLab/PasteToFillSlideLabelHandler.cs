using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label.PasteLab
{
    [ExportLabelRibbonId(
        "PasteToFillSlideMenuFrame",
        "PasteToFillSlideMenuShape",
        "PasteToFillSlideMenuLine",
        "PasteToFillSlideMenuFreeform",
        "PasteToFillSlideMenuPicture",
        "PasteToFillSlideMenuGroup",
        "PasteToFillSlideMenuInk",
        "PasteToFillSlideMenuVideo",
        "PasteToFillSlideMenuTextEdit",
        "PasteToFillSlideMenuChart",
        "PasteToFillSlideMenuTable",
        "PasteToFillSlideMenuTableWhole",
        "PasteToFillSlideMenuSmartArtBackground",
        "PasteToFillSlideMenuSmartArtEditSmartArt",
        "PasteToFillSlideMenuSmartArtEditText")]
    class PasteToFillSlideLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.PasteLabText.PasteToFillSlide;
        }
    }
}