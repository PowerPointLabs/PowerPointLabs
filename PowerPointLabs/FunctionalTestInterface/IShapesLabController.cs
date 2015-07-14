using System.Collections.Generic;

namespace FunctionalTestInterface
{
    public interface IShapesLabController
    {
        void OpenPane();
        void SaveSelectedShapes();
        IShapesLabLabeledThumbnail GetLabeledThumbnail(string labelName);
        void ImportLibrary(string pathToLibrary);
        void ImportShape(string pathToShape);
        List<ISlideData> FetchShapeGalleryPresentationData();
    }
}
