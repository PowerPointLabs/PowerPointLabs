using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Views;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId(
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
    class EditNameActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var selection = this.GetCurrentSelection();
            var selectedShape = selection.ShapeRange[1];
            if (selection.HasChildShapeRange)
            {
                selectedShape = selection.ChildShapeRange[1];
            }
            
            var editForm = new EditNameDialogBox(selectedShape);
            editForm.ShowDialog();
        }
    }
}
