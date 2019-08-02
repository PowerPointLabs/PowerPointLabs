using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ActionFramework.ShortcutsLab
{
    [ExportActionRibbonId(ShortcutsLabText.HideShapeTag)]
    class HideShapeActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            Selection selection = this.GetCurrentSelection();
            ShapeRange selectedShapes = ShapeUtil.GetShapeRange(selection);

            selectedShapes.Visible = Microsoft.Office.Core.MsoTriState.msoFalse;
        }
    }
}
