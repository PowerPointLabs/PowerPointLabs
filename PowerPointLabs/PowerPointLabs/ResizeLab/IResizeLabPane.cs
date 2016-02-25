using System;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ResizeLab
{
    public delegate void Action(PowerPoint.ShapeRange selectedShapes);
    public delegate Action ResizeAction(PowerPoint.ShapeRange selectedShapes);

    public interface IResizeLabPane
    {
        void ShowErrorMessageBox(string content, Exception exception = null);
        void Preview(PowerPoint.ShapeRange selectedShapes, ResizeAction previewAction);
        void Reset();
    }
}