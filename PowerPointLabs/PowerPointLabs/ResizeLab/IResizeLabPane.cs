using System;

using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{

    public interface IResizeLabPane
    {
        void ShowErrorMessageBox(string content, Exception exception = null);
        void Preview(PowerPoint.ShapeRange selectedShapes, Action<PowerPoint.ShapeRange> previewAction, int minNumberofSelectedShapes);
        void Preview(PowerPoint.ShapeRange selectedShapes, float referenceWidth, float referenceHeight, Action<PowerPoint.ShapeRange, float, float, bool> previewAction);
        void Reset();
        void ExecuteResizeAction(PowerPoint.ShapeRange selectedShapes, Action<PowerPoint.ShapeRange> resizeAction);
        void ExecuteResizeAction(PowerPoint.ShapeRange selectedShapes, float slideWidth, float slideHeight,
            Action<PowerPoint.ShapeRange, float, float, bool> resizeAction);
    }
}