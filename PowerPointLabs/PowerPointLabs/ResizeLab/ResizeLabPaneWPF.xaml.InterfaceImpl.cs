using System;
using System.Collections.Generic;
using System.Windows;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Utils;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    public partial class ResizeLabPaneWPF
    {
        private readonly Dictionary<string, ShapeProperties> _originalShapeProperties = new Dictionary<string, ShapeProperties>();

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


        public void Preview(PowerPoint.ShapeRange selectedShapes, Action<PowerPoint.ShapeRange> previewAction, int minNumberofSelectedShapes)
        {
            if (selectedShapes == null || selectedShapes.Count < minNumberofSelectedShapes) return;

            StoreOriginalShapesProperties(selectedShapes);
            previewAction.Invoke(selectedShapes);
        }

        public void Preview(PowerPoint.ShapeRange selectedShapes, float referenceWidth, float referenceHeight, Action<PowerPoint.ShapeRange, float, float, bool> previewAction)
        {
            if (selectedShapes == null) return;

            StoreOriginalShapesProperties(selectedShapes);
            previewAction.Invoke(selectedShapes, referenceWidth, referenceHeight, IsAspectRatioLocked);
        }

        public void Reset()
        {
            var selectedShapes = GetSelectedShapes(false);

            if (selectedShapes != null)
            {
                _resizeLab.ResetShapes(selectedShapes, _originalShapeProperties);
            }
        }

        public void ExecuteResizeAction(PowerPoint.ShapeRange selectedShapes, Action<PowerPoint.ShapeRange> resizeAction)
        {
            if (selectedShapes == null) return;

            Reset();
            resizeAction.Invoke(selectedShapes);
            CleanOriginalShapes();
        }

        public void ExecuteResizeAction(PowerPoint.ShapeRange selectedShapes, float slideWidth, float slideHeight, Action<PowerPoint.ShapeRange, float, float, bool> resizeAction)
        {
            if (selectedShapes == null) return;

            Reset();
            resizeAction.Invoke(selectedShapes, slideWidth, slideHeight, IsAspectRatioLocked);
            CleanOriginalShapes();
        }

        private void StoreOriginalShapesProperties(PowerPoint.ShapeRange selectedShapes)
        {
            _originalShapeProperties.Clear();

            for (int i = 1; i <= selectedShapes.Count; i++)
            {
                var shape = new PPShape(selectedShapes[i]);
                var properties = new ShapeProperties(shape.Name, shape.Top, shape.Left, shape.AbsoluteWidth, shape.AbsoluteHeight, shape.ShapeRotation);
                _originalShapeProperties.Add(shape.Name, properties);
            }
        }

        private void CleanOriginalShapes()
        {
            _originalShapeProperties.Clear();
        }
    }
}
