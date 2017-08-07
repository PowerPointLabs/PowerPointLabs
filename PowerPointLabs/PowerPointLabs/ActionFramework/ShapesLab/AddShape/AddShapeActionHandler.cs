using System.IO;
using System.Windows.Forms;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ShortcutsLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShapesLab
{
    [ExportActionRibbonId(ShortcutsLabText.AddCustomShapeTag)]
    class AddShapeActionHandler : ShapesLabActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            var customShape = InitCustomShapePane();
            var selection = this.GetCurrentSelection();
            var addIn = this.GetAddIn();
            // first of all we check if the shape gallery has been opened correctly
            if (!addIn.ShapePresentation.Opened)
            {
                MessageBox.Show(CommonText.ErrorShapeGalleryInit);
                return;
            }

            var selectedShapes = selection.ShapeRange;
            if (selection.HasChildShapeRange)
            {
                selectedShapes = selection.ChildShapeRange;
            }

            // add shape into shape gallery first to reduce flicker
            var shapeName = addIn.ShapePresentation.AddShape(selectedShapes, selectedShapes[1].Name);

            // add the selection into pane and save it as .png locally
            var shapeFullName = Path.Combine(customShape.CurrentShapeFolderPath, shapeName + ".png");
            ConvertToPicture.ConvertAndSave(selection, shapeFullName);

            // sync the shape among all opening panels
            addIn.SyncShapeAdd(shapeName, shapeFullName, customShape.CurrentCategory);

            // finally, add the shape into the panel and waiting for name editing
            customShape.AddCustomShape(shapeName, shapeFullName, true);

            SetPaneVisibility(true);
        }
    }
}
