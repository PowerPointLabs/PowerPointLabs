using System.IO;
using System.Windows.Forms;

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
            CustomShapePane customShape = InitCustomShapePane();
            Selection selection = this.GetCurrentSelection();
            ThisAddIn addIn = this.GetAddIn();
            // first of all we check if the shape gallery has been opened correctly
            if (!addIn.ShapePresentation.Opened)
            {
                MessageBox.Show(CommonText.ErrorShapeGalleryInit);
                return;
            }
            customShape.AddCustomShapeToPane(selection, addIn);

            SetPaneVisibility(true);
        }
    }
}
