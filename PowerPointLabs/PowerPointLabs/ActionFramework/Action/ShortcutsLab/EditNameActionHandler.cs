using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Views;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId(
        TextCollection.EditNameId + TextCollection.MenuShape,
        TextCollection.EditNameId + TextCollection.MenuLine,
        TextCollection.EditNameId + TextCollection.MenuFreeform,
        TextCollection.EditNameId + TextCollection.MenuPicture,
        TextCollection.EditNameId + TextCollection.MenuGroup,
        TextCollection.EditNameId + TextCollection.MenuInk,
        TextCollection.EditNameId + TextCollection.MenuVideo,
        TextCollection.EditNameId + TextCollection.MenuTextEdit,
        TextCollection.EditNameId + TextCollection.MenuChart,
        TextCollection.EditNameId + TextCollection.MenuTable,
        TextCollection.EditNameId + TextCollection.MenuTableCell,
        TextCollection.EditNameId + TextCollection.MenuSmartArt,
        TextCollection.EditNameId + TextCollection.MenuEditSmartArt,
        TextCollection.EditNameId + TextCollection.MenuEditSmartArtText)]
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
