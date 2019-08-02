using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.ShortcutsLab.Views;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportActionRibbonId(ShortcutsLabText.EditNameTag)]
    class EditNameActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            Selection selection = this.GetCurrentSelection();
            Shape selectedShape = ShapeUtil.GetShapeRange(selection)[1];
            
            EditNameDialogBox editForm = new EditNameDialogBox(selectedShape);
            editForm.ShowThematicDialog();

            if (!this.GetApplication().CommandBars.GetPressedMso("SelectionPane"))
            {
                this.GetApplication().CommandBars.ExecuteMso("SelectionPane");
            }
        }
    }
}
