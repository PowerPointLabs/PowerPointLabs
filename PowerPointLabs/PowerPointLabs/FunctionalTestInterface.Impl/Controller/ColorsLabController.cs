using System;
using System.Collections.Generic;
using System.Drawing;
using System.Windows;
using System.Windows.Controls.Primitives;
using System.Windows.Input;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ColorsLab;
using PowerPointLabs.TextCollection;

using TestInterface;

using Point = System.Windows.Point;
using Rectangle = System.Windows.Shapes.Rectangle;

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
                    LeftClickOnRectangle(rect);

                    
                }));
            }
        }

        public void ClickAnalogousRect(int index)
        {
            System.Windows.Shapes.Rectangle rect;

            switch (index)
            {
                case 1:
                rect = _pane.ColorsLabPaneWPF1.analogousLowerRectangle;
                break;
                case 2:
                rect = _pane.ColorsLabPaneWPF1.analogousMiddleRectangle;
                break;
                case 3:
                rect = _pane.ColorsLabPaneWPF1.analogousHigherRectangle;
                break;
                default:
                rect = null;
                break;
            }

            if (_pane != null && rect != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    LeftClickOnRectangle(rect);
                }));
            }
        }

        public void ClickComplementaryRect(int index)
        {
            System.Windows.Shapes.Rectangle rect;

            switch (index)
            {
                case 1:
                rect = _pane.ColorsLabPaneWPF1.complementaryLowerRectangle;
                break;
                case 2:
                rect = _pane.ColorsLabPaneWPF1.complementaryMiddleRectangle;
                break;
                case 3:
                rect = _pane.ColorsLabPaneWPF1.complementaryHigherRectangle;
                break;
                default:
                rect = null;
                break;
            }

            if (_pane != null && rect != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    LeftClickOnRectangle(rect);
                }));
            }
        }

        public void ClickTriadicRect(int index)
        {
            System.Windows.Shapes.Rectangle rect;

            switch (index)
            {
                case 1:
                rect = _pane.ColorsLabPaneWPF1.triadicLowerRectangle;
                break;
                case 2:
                rect = _pane.ColorsLabPaneWPF1.triadicMiddleRectangle;
                break;
                case 3:
                rect = _pane.ColorsLabPaneWPF1.triadicHigherRectangle;
                break;
                default:
                rect = null;
                break;
            }

            if (_pane != null && rect != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    LeftClickOnRectangle(rect);
                }));
            }
        }

        public void ClickTetradicRect(int index)
        {
            System.Windows.Shapes.Rectangle rect;

            switch (index)
            {
                case 1:
                rect = _pane.ColorsLabPaneWPF1.tetradicOneRectangle;
                break;
                case 2:
                rect = _pane.ColorsLabPaneWPF1.tetradicTwoRectangle;
                break;
                case 3:
                rect = _pane.ColorsLabPaneWPF1.tetradicThreeRectangle;
                break;
                case 4:
                rect = _pane.ColorsLabPaneWPF1.tetradicFourRectangle;
                break;
                default:
                rect = null;
                break;
            }

            if (_pane != null && rect != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    LeftClickOnRectangle(rect);
                }));
            }
        }

        public void LoadFavoriteColors(string filePath)
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    _pane.ColorsLabPaneWPF1.LoadFavoriteColorsFromPath(filePath);
                }));
            }
        }

        public void ResetFavoriteColors()
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    _pane.ColorsLabPaneWPF1.reloadColorButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                }));
            }
        }

        public void ClearFavoriteColors()
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    _pane.ColorsLabPaneWPF1.clearColorButton.RaiseEvent(new RoutedEventArgs(ButtonBase.ClickEvent));
                }));
            }
        }

        public void ClearRecentColors()
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    _pane.ColorsLabPaneWPF1.EmptyRecentColorsPanel();
                }));
            }
        }

        public IList<Color> GetCurrentFavoritePanel()
        {
            IList<Color> list = null;
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    list = _pane.ColorsLabPaneWPF1.GetFavoriteColorsPanelAsList();
                }));
            }
            return list;
        }

        public IList<Color> GetCurrentRecentPanel()
        {
            IList<Color> list = null;
            if (_pane != null)
            {
                UIThreadExecutor.Execute((Action)(() =>
                {
                    list = _pane.ColorsLabPaneWPF1.GetRecentColorsPanelAsList();
                }));
            }
            return list;
        }

        private void LeftClickOnRectangle(Rectangle rect)
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
        }
    }
}
