using System;
using System.Windows;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    public partial class ResizeLabPaneWPF
    {
        public void ShowErrorMessageBox(string content, Exception exception = null)
        {
            if (exception != null)
            {
                Views.ErrorDialogWrapper.ShowDialog("Error", content, exception);
            }
            else
            {
                MessageBox.Show(content, "Error");
            }
        }

        public void Preview(PowerPoint.ShapeRange selectedShapes, SingleInputResizeAction previewAction, int minNoOfSelectedShapes)
        {
            if (selectedShapes == null || selectedShapes.Count < minNoOfSelectedShapes) return;

            var action = previewAction(selectedShapes);

            StoreOriginalShapesProperties(selectedShapes);
            action(selectedShapes);
        }

        public void Preview(PowerPoint.ShapeRange selectedShapes, float slideWidth, float slideHeight,
            MultiInputResizeAction previewAction)
        {
            if (selectedShapes == null) return;

            var action = previewAction(selectedShapes, slideWidth, slideHeight, IsAspectRatioLocked);

            StoreOriginalShapesProperties(selectedShapes);
            action(selectedShapes, slideWidth, slideHeight, IsAspectRatioLocked);
        }

        public void Reset()
        {
            var selectedShapes = GetSelectedShapes(false);

            if (selectedShapes != null)
            {
                _resizeLab.ResetShapes(selectedShapes, _originalShapeProperties);
            }
        }

        public void ExecuteResizeAction(PowerPoint.ShapeRange selectedShapes, SingleInputResizeAction resizeAction)
        {
            if (selectedShapes == null) return;

            var action = resizeAction(selectedShapes);

            Reset();
            action(selectedShapes);
            CleanOriginalShapes();
        }

        public void ExecuteResizeAction(PowerPoint.ShapeRange selectedShapes, float slideWidth, float slideHeight, MultiInputResizeAction resizeAction)
        {
            if (selectedShapes == null) return;

            var action = resizeAction(selectedShapes, slideWidth, slideHeight, IsAspectRatioLocked);

            Reset();
            action(selectedShapes, slideWidth, slideHeight, IsAspectRatioLocked);
            CleanOriginalShapes();
        }

        private void StoreOriginalShapesProperties(PowerPoint.ShapeRange selectedShapes)
        {
            _originalShapeProperties.Clear();

            for (int i = 1; i <= selectedShapes.Count; i++)
            {
                var shape = new PPShape(selectedShapes[i]);
                var properties = new ShapeProperties(shape.Name, shape.Top, shape.Left, shape.AbsoluteWidth, shape.AbsoluteHeight);
                _originalShapeProperties.Add(shape.Name, properties);
            }
        }
    }
}
