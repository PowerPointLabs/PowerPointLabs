using System;
using System.Windows;
using System.Windows.Input;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ColorsLab;
using PowerPointLabs.TextCollection;

using TestInterface;

using Button = System.Windows.Controls.Button;

namespace PowerPointLabs.FunctionalTestInterface.Impl.Controller
{
    [Serializable]
    class ColorsLabController : MarshalByRefObject, IColorsLabController
    {
        private static IColorsLabController _instance = new ColorsLabController();

        public static IColorsLabController Instance { get { return _instance; } }

        private ColorsLabPane _pane;

        private ColorsLabController() {}

        public void OpenPane()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                FunctionalTestExtensions.GetRibbonUi().OnAction(
                    new RibbonControl(ColorsLabText.PaneTag));
                _pane = FunctionalTestExtensions.GetTaskPane(
                    typeof(ColorsLabPane)).Control as ColorsLabPane;
            }));
        }

        public Point GetApplyTextButtonLocation()
        {
            Point point = new Point(0, 0);
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    point = _pane.ColorsLabPaneWPF1.GetApplyTextButtonLocationAsPoint();
                }));
            }
            return point;
        }

        public Point GetApplyLineButtonLocation()
        {
            Point point = new Point(0, 0);
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    point = _pane.ColorsLabPaneWPF1.GetApplyLineButtonLocationAsPoint();
                }));
            }
            return point;
        }

        public Point GetApplyFillButtonLocation()
        {
            Point point = new Point(0, 0);
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    point = _pane.ColorsLabPaneWPF1.GetApplyFillButtonLocationAsPoint();
                }));
            }
            return point;
        }

        public Point GetMainColorRectangleLocation()
        {
            Point point = new Point(0, 0);
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    point = _pane.ColorsLabPaneWPF1.GetMainColorRectangleLocationAsPoint();
                }));
            }
            return point;
        }

        public Point GetEyeDropperButtonLocation()
        {
            Point point = new Point(0, 0);
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    point = _pane.ColorsLabPaneWPF1.GetEyeDropperButtonLocationAsPoint();
                }));
            }
            return point;
        }

        public void SlideBrightnessSlider(int value)
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    _pane.ColorsLabPaneWPF1.brightnessSlider.Value = value;
                }));
            }
        }

        public void SlideSaturationSlider(int value)
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    _pane.ColorsLabPaneWPF1.saturationSlider.Value = value;
                }));
            }
        }

        public void ClickMonochromeRect(int index)
        {
            System.Windows.Shapes.Rectangle rect;

            switch (index)
            {
                case 1:
                    rect = _pane.ColorsLabPaneWPF1.monochromaticRectangleOne;
                    break;
                case 2:
                    rect = _pane.ColorsLabPaneWPF1.monochromaticRectangleTwo;
                    break;
                case 3:
                    rect = _pane.ColorsLabPaneWPF1.monochromaticRectangleThree;
                    break;
                case 4:
                    rect = _pane.ColorsLabPaneWPF1.monochromaticRectangleFour;
                    break;
                case 5:
                    rect = _pane.ColorsLabPaneWPF1.monochromaticRectangleFive;
                    break;
                case 6:
                    rect = _pane.ColorsLabPaneWPF1.monochromaticRectangleSix;
                    break;
                case 7:
                    rect = _pane.ColorsLabPaneWPF1.monochromaticRectangleSeven;
                    break;
                default:
                    rect = null;
                    break;
            }

            if (_pane != null && rect != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    rect.RaiseEvent(
                        new MouseButtonEventArgs(Mouse.PrimaryDevice, Environment.TickCount, MouseButton.Left)
                        {
                            RoutedEvent = UIElement.MouseDownEvent
                        });

                    rect.RaiseEvent(
                       new MouseButtonEventArgs(Mouse.PrimaryDevice, Environment.TickCount, MouseButton.Left)
                       {
                           RoutedEvent = UIElement.MouseUpEvent
                       });
                }));
            }
        }

    }
}
