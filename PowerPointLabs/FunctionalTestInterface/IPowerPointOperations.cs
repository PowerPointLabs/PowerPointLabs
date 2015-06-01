using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
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
    }
}
