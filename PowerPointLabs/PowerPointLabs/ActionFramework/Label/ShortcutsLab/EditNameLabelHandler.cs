using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Label
{
    [ExportLabelRibbonId(
        TextCollection.EditNameMenuId + TextCollection.MenuShape,
        TextCollection.EditNameMenuId + TextCollection.MenuLine,
        TextCollection.EditNameMenuId + TextCollection.MenuFreeform,
        TextCollection.EditNameMenuId + TextCollection.MenuPicture,
        TextCollection.EditNameMenuId + TextCollection.MenuGroup,
        TextCollection.EditNameMenuId + TextCollection.MenuInk,
        TextCollection.EditNameMenuId + TextCollection.MenuVideo,
        TextCollection.EditNameMenuId + TextCollection.MenuTextEdit,
        TextCollection.EditNameMenuId + TextCollection.MenuChart,
        TextCollection.EditNameMenuId + TextCollection.MenuTable,
        TextCollection.EditNameMenuId + TextCollection.MenuTableCell,
        TextCollection.EditNameMenuId + TextCollection.MenuSmartArt,
        TextCollection.EditNameMenuId + TextCollection.MenuEditSmartArt,
        TextCollection.EditNameMenuId + TextCollection.MenuEditSmartArtText)]
    class EditNameLabelHandler : LabelHandler
    {
        protected override string GetLabel(string ribbonId)
        {
            return TextCollection.EditNameShapeLabel;
        }
    }
}
