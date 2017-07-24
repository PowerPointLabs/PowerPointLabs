using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportActionRibbonId(ShortcutsLabText.HideShapeTag)]
    class HideShapeActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var selection = this.GetCurrentSelection();
            var selectedShapes = selection.ShapeRange;
            if (selection.HasChildShapeRange)
            {
                selectedShapes = selection.ChildShapeRange;
            }

            selectedShapes.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
        }
    }
}
