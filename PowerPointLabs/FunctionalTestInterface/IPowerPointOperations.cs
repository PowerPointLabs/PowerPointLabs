using System.IO;
using Microsoft.Office.Interop.PowerPoint;

namespace FunctionalTestInterface
{
    public interface IPowerPointOperations
    {
        void EnterFunctionalTest();
        void ExitFunctionalTest();
        bool IsInFunctionalTest();

        Slide GetCurrentSlide();
        Slide SelectSlide(int index);
        Slide SelectSlide(string slideName);

        Selection GetCurrentSelection();
        ShapeRange SelectShapes(string shapeName);
        ShapeRange SelectShapesByPrefix(string prefix);
        FileInfo ExportSelectedShapes();
    }
}
