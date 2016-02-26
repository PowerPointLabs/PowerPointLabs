using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    public delegate void SingleInputAction(PowerPoint.ShapeRange selectedShapes);
    public delegate SingleInputAction SingleInputResizeAction(PowerPoint.ShapeRange selectedShapes);

    public delegate void MultiInputAction(
        PowerPoint.ShapeRange selectedShapes, float slideWidth, float slideHeight, bool isAspectRatio);
    public delegate MultiInputAction MultiInputResizeAction(
        PowerPoint.ShapeRange selectedShapes, float slideWidth, float slideHeight, bool isAspectRatio);

    public interface IResizeLabPane
    {
        void ShowErrorMessageBox(string content, Exception exception = null);
        void Preview(PowerPoint.ShapeRange selectedShapes, SingleInputResizeAction previewAction);
        void Preview(PowerPoint.ShapeRange selectedShapes, float slideWidth, float slideHeight,
            MultiInputResizeAction previewAction);
        void Reset();
        void ExecuteResizeAction(PowerPoint.ShapeRange selectedShapes, SingleInputResizeAction resizeAction);
        void ExecuteResizeAction(PowerPoint.ShapeRange selectedShapes, float slideWidth, float slideHeight,
            MultiInputResizeAction resizeAction);
    }
}