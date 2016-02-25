using System;
using System.Collections.Generic;
using System.Windows;
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

        public void Preview(PowerPoint.ShapeRange selectedShapes, SingleInputResizeAction previewAction)
        {
            if (selectedShapes == null) return;

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

        private void StoreOriginalShapesProperties(PowerPoint.ShapeRange selectedShapes)
        {
            _originalShapeProperties.Clear();

            for (int i = 1; i <= selectedShapes.Count; i++)
            {
                var shape = selectedShapes[i];
                var properties = new ShapeProperties(shape.Name, shape.Top, shape.Left, shape.Width, shape.Height);
                _originalShapeProperties.Add(shape.Name, properties);
            }
        }
    }
}
