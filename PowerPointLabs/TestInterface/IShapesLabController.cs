using System.Collections.Generic;
using System.Windows;

namespace TestInterface
{
    public interface IShapesLabController
    {
        void OpenPane();
        void SaveSelectedShapes();
        Point GetShapeForClicking(string shapeName);
        void ImportLibrary(string pathToLibrary);
        void ImportShape(string pathToShape);
        List<ISlideData> FetchShapeGalleryPresentationData();
        void ClickAddShapeButton();
        bool GetAddShapeButtonStatus();
    }
}
