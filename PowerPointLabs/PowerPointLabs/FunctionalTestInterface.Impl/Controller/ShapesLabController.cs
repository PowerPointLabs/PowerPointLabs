using System;
using System.Collections.Generic;
using System.Linq;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ShapesLab;
using PowerPointLabs.TextCollection;

using TestInterface;


namespace PowerPointLabs.FunctionalTestInterface.Impl.Controller
{
    [Serializable]
    class ShapesLabController : MarshalByRefObject, IShapesLabController
    {
        private static IShapesLabController _instance = new ShapesLabController();

        public static IShapesLabController Instance { get { return _instance; } }

        private CustomShapePane_ _pane;

        private ShapesLabController() {}

        public void OpenPane()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                FunctionalTestExtensions.GetRibbonUi().OnAction(
                    new RibbonControl(ShapesLabText.PaneTag));
                _pane = FunctionalTestExtensions.GetTaskPane(
                    typeof(CustomShapePane_)).Control as CustomShapePane_;
            }));
        }

        public void SaveSelectedShapes()
        {
            UIThreadExecutor.Execute(() =>
            {
                FunctionalTestExtensions.GetRibbonUi().OnAction(
                    new RibbonControl("AddShape"));
            });
        }

        public IShapesLabLabeledThumbnail GetLabeledThumbnail(string labelName)
        {
            if (_pane != null)
            {
                return _pane.GetLabeledThumbnail(labelName);
            }
            return null;
        }

        public void ImportLibrary(string pathToLibrary)
        {
            if (_pane != null)
            {
                _pane.ImportLibrary(pathToLibrary);
            }
        }

        public void ImportShape(string pathToShape)
        {
            if (_pane != null)
            {
                _pane.ImportShape(pathToShape);
            }
        }

        public List<ISlideData> FetchShapeGalleryPresentationData()
        {
            if (_pane != null)
            {
                List<ISlideData> slideData = _pane.GetShapeGallery()
                    .Slides.Cast<Slide>().Select(SlideData.FromSlide).ToList();
                return slideData;
            }
            return null;
        }

        public void ClickAddShapeButton()
        {
            if (_pane != null && _pane.GetAddShapeButton() != null)
            {
                // Perform clicking of button on its own UI thread
                UIThreadExecutor.Execute(() =>
                {
                    _pane.GetAddShapeButton().PerformClick();
                });
            }

        }

        public bool GetAddShapeButtonStatus()
        {
            return _pane.GetAddShapeButton().Enabled;
        }
    }
}
