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

            var duplicatedShapes = selectedShapes.Duplicate();
            var action = previewAction(selectedShapes);

            SetOriginalTopLeft(selectedShapes, duplicatedShapes);
            StoreOriginalShapes(duplicatedShapes);

            action(selectedShapes);
        }

        public void Reset()
        {
            var selectedShapes = GetSelectedShapes(false);
            if (selectedShapes != null)
            {
                _resizeLab.ResetShapes(selectedShapes, _originalShapes);
            }
        }

        /// <summary>
        /// As the properties of top and left does not maintain after duplication,
        /// this method enforces the same properties of top and left.
        /// </summary>
        /// <param name="selectedShapes"></param>
        /// <param name="duplicatedShapes"></param>
        private static void SetOriginalTopLeft(PowerPoint.ShapeRange selectedShapes, PowerPoint.ShapeRange duplicatedShapes)
        {
            for (int i = 1; i <= selectedShapes.Count; i++)
            {
                duplicatedShapes[i].Top = selectedShapes[i].Top;
                duplicatedShapes[i].Left = selectedShapes[i].Left;
            }
        }

        /// <summary>
        /// Store the properties of the original shapes to the dictionary.
        /// </summary>
        /// <param name="selectedShapes"></param>
        private void StoreOriginalShapes(PowerPoint.ShapeRange selectedShapes)
        {
            _originalShapes.Clear();

            for (int i = 1; i <= selectedShapes.Count; i++)
            {
                var shape = selectedShapes[i];
                var shapeName = shape.Name;
                _originalShapes.Add(shapeName, shape);
            }
        }
    }
}
