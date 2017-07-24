using System;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.PositionsLab;
using PowerPointLabs.TextCollection;

using TestInterface;

namespace PowerPointLabs.FunctionalTestInterface.Impl.Controller
{
    [Serializable]
    class PositionsLabController : MarshalByRefObject, IPositionsLabController
    {
        private static IPositionsLabController _instance = new PositionsLabController();

        public static IPositionsLabController Instance { get { return _instance; } }

        private PositionsPane _pane;

        private PositionsLabController() { }

        public void OpenPane()
        {
            UIThreadExecutor.Execute((Action)(() =>
            {
                FunctionalTestExtensions.GetRibbonUi().OnAction(
                    new RibbonControl(PositionsLabText.PaneTag));
                _pane = FunctionalTestExtensions.GetTaskPane(
                    typeof(PositionsPane)).Control as PositionsPane;
            }));
        }

        public void ToggleRotateButton()
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    _pane.positionsPaneWPF.rotationButton.IsChecked = !_pane.positionsPaneWPF.rotationButton.IsChecked;
                });
            }
        }

        public void ToggleDuplicateAndRotateButton()
        {
            if (_pane != null)
            {
                UIThreadExecutor.Execute(() =>
                {
                    _pane.positionsPaneWPF.duplicateRotationButton.IsChecked = !_pane.positionsPaneWPF.duplicateRotationButton.IsChecked;
                });
            }
        }

        public void ReorientFixed()
        {
            PositionsLabSettings.ReorientShapeOrientation = PositionsLabSettings.RadialShapeOrientationObject.Fixed;
        }

        public void ReorientDynamic()
        {
            PositionsLabSettings.ReorientShapeOrientation = PositionsLabSettings.RadialShapeOrientationObject.Dynamic;
        }
    }
}
