using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.Views;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId(TextCollection.EditNameTag)]
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
