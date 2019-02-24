
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ShapesLab;

using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShapesLab
{
    [ExportActionRibbonId(ShortcutsLabText.AddCustomShapeTag)]
    class AddShapeActionHandler : ShapesLabActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            //CustomShapePane_ customShape = InitCustomShapePane();
            //Selection selection = this.GetCurrentSelection();
            //ThisAddIn addIn = this.GetAddIn();

            //customShape.AddShapeFromSelection(selection, addIn);
        }
    }
}
