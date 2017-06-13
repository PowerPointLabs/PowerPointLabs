using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        "ConvertToPictureMenuShape",
        "ConvertToPictureMenuLine",
        "ConvertToPictureMenuFreeform",
        "ConvertToPictureMenuGroup",
        "ConvertToPictureMenuInk",
        "ConvertToPictureMenuVideo",
        "ConvertToPictureMenuChart",
        "ConvertToPictureMenuTable",
        "ConvertToPictureMenuTableWhole",
        "ConvertToPictureMenuSmartArtBackground",
        "ConvertToPictureMenuSmartArtEditSmartArt",
        "ConvertToPictureMenuSmartArtEditText")]
    class ConvertToPictureShapeLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.ConvertToPictureShapeLabel;
        }
    }
}
