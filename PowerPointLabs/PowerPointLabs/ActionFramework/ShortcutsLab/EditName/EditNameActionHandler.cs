using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ShortcutsLab.Views;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportActionRibbonId(ShortcutsLabText.EditNameTag)]
    class EditNameActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            Selection selection = this.GetCurrentSelection();
            Shape selectedShape = selection.ShapeRange[1];
            if (selection.HasChildShapeRange)
            {
                selectedShape = selection.ChildShapeRange[1];
            }
            
            EditNameDialogBox editForm = new EditNameDialogBox(selectedShape);
            editForm.ShowDialog();

            if (!this.GetApplication().CommandBars.GetPressedMso("SelectionPane"))
            {
                this.GetApplication().CommandBars.ExecuteMso("SelectionPane");
            }
        }
    }
}
