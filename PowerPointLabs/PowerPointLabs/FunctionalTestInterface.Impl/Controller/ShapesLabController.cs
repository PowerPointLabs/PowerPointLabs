using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls.Primitives;

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

        private CustomShapePane _pane;

        private ShapesLabController() {}

        public void OpenPane()
        {
            UIThreadExecutor.Execute(() =>
            {
                FunctionalTestExtensions.GetRibbonUi().OnAction(
                    new RibbonControl(ShapesLabText.PaneTag));
                _pane = FunctionalTestExtensions.GetTaskPane(
                    typeof(CustomShapePane)).Control as CustomShapePane;
            });
            _pane?.InitCustomShapePaneStorage();
        }

        public void SaveSelectedShapes()
        {
            UIThreadExecutor.Execute(() =>
            {
                _pane?.SaveSelectedShapes();
            });
        }

        public System.Windows.Point GetShapeForClicking(string shapeName)
        {
            System.Windows.Point point = new System.Windows.Point(0, 0);
            if (_pane == null)
            {
                return point;
            }
            Task task = Task.Factory.StartNew(() =>
            {
                _pane.CustomShapePaneWPF1.Dispatcher.Invoke(() =>
                {
                    point = _pane.GetShapeForClicking(shapeName);
                });
            });
            task.Wait();
            return point;
        }

        public void ImportLibrary(string pathToLibrary)
        {
            if (_pane == null)
            {
                return;
            }
            Task task = Task.Factory.StartNew(() =>
            {
                _pane.CustomShapePaneWPF1.Dispatcher.Invoke(() =>
                {
                    _pane.ImportLibrary(pathToLibrary);
                });
            });
            task.Wait();
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
            if (_pane == null)
            {
                return null;
            }
            List<ISlideData> slideData = _pane.GetShapeGallery()
                .Slides.Cast<Slide>().Select(SlideData.FromSlide).ToList();
            return slideData;
        }

        public void ClickAddShapeButton()
        {
            if (_pane == null || _pane.GetAddShapeButton() == null)
            {
                return;
            }
            // Perform clicking of button on its own UI thread
            UIThreadExecutor.Execute(() =>
            {
                _pane.GetAddShapeButton().RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
            });

        }

        public bool GetAddShapeButtonStatus()
        {
            bool addShapeButtonStatus = false;
            Task task = Task.Factory.StartNew(() =>
            {
                _pane.GetAddShapeButton().Dispatcher.Invoke(() =>
                {
                    addShapeButtonStatus = _pane.GetAddShapeButton().IsEnabled;
                });
            });
            task.Wait();
            return addShapeButtonStatus;
        }
    }
}
