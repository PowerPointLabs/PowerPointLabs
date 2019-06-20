using System;
using System.Windows;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.Utils;
using PowerPointLabs.Views;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    public partial class ResizeLabPaneWPF
    {
        private bool _isPreview;

        public void ShowErrorMessageBox(string content, Exception exception = null)
        {
            if (exception != null)
            {
                ErrorDialogBox.ShowDialog(TextCollection.CommonText.ErrorTitle, content, exception);
            }
            else
            {
                MessageBoxUtil.Show(content, TextCollection.CommonText.ErrorTitle);
            }
        }


        public void Preview(PowerPoint.ShapeRange selectedShapes, Action<PowerPoint.ShapeRange> previewAction, int minNumberofSelectedShapes)
        {
            if (selectedShapes == null || selectedShapes.Count < minNumberofSelectedShapes)
            {
                return;
            }

            _isPreview = true;
            this.StartNewUndoEntry();
            previewAction.Invoke(selectedShapes);
        }

        public void Preview(PowerPoint.ShapeRange selectedShapes, float referenceWidth, float referenceHeight, Action<PowerPoint.ShapeRange, float, float, bool> previewAction)
        {
            if (selectedShapes == null)
            {
                return;
            }

            _isPreview = true;
            this.StartNewUndoEntry();
            previewAction.Invoke(selectedShapes, referenceWidth, referenceHeight, IsAspectRatioLocked);
        }

        public void Reset()
        {
            PowerPoint.ShapeRange selectedShapes = GetSelectedShapes(false);

            if (selectedShapes != null && _isPreview)
            {
                this.ExecuteOfficeCommand("Undo");
                GC.Collect();
                _isPreview = false;
            }
        }

        public void ExecuteResizeAction(PowerPoint.ShapeRange selectedShapes, Action<PowerPoint.ShapeRange> resizeAction)
        {
            if (selectedShapes == null)
            {
                return;
            }

            Reset();
            this.StartNewUndoEntry();
            resizeAction.Invoke(selectedShapes);
            _isPreview = false;
        }

        public void ExecuteResizeAction(PowerPoint.ShapeRange selectedShapes, float slideWidth, float slideHeight, Action<PowerPoint.ShapeRange, float, float, bool> resizeAction)
        {
            if (selectedShapes == null)
            {
                return;
            }

            Reset();
            this.StartNewUndoEntry();
            resizeAction.Invoke(selectedShapes, slideWidth, slideHeight, IsAspectRatioLocked);
            _isPreview = false;
        }
    }
}
