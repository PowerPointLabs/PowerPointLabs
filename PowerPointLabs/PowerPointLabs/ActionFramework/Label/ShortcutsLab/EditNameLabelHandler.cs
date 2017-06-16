using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        TextCollection.EditNameTag + TextCollection.MenuShape,
        TextCollection.EditNameTag + TextCollection.MenuLine,
        TextCollection.EditNameTag + TextCollection.MenuFreeform,
        TextCollection.EditNameTag + TextCollection.MenuPicture,
        TextCollection.EditNameTag + TextCollection.MenuGroup,
        TextCollection.EditNameTag + TextCollection.MenuInk,
        TextCollection.EditNameTag + TextCollection.MenuVideo,
        TextCollection.EditNameTag + TextCollection.MenuTextEdit,
        TextCollection.EditNameTag + TextCollection.MenuChart,
        TextCollection.EditNameTag + TextCollection.MenuTable,
        TextCollection.EditNameTag + TextCollection.MenuTableCell,
        TextCollection.EditNameTag + TextCollection.MenuSmartArt,
        TextCollection.EditNameTag + TextCollection.MenuEditSmartArt,
        TextCollection.EditNameTag + TextCollection.MenuEditSmartArtText)]
    class EditNameLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.EditNameShapeLabel;
        }
    }
}
