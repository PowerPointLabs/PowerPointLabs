using System.Collections.Generic;
using System.Windows.Forms;

namespace TestInterface
{
    public interface IShapesLabController
    {
        void OpenPane();
        void SaveSelectedShapes();
        IShapesLabLabeledThumbnail GetLabeledThumbnail(string labelName);
        void ImportLibrary(string pathToLibrary);
        void ImportShape(string pathToShape);
        List<ISlideData> FetchShapeGalleryPresentationData();
        void ClickAddShapeButton();
    }
}
