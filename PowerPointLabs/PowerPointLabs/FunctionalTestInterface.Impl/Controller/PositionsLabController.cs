using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPointLabs.ActionFramework.Common.Extension;
using TestInterface;
using PowerPointLabs.PositionsLab;

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
            UIThreadExecutor.Execute(() =>
            {
                FunctionalTestExtensions.GetRibbonUi().OnAction(
                    new RibbonControl("PositionsLabButton"));
                _pane = FunctionalTestExtensions.GetTaskPane(
                    typeof(PositionsPane)).Control as PositionsPane;
            });
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
            PositionsLabMain.ReorientShapeOrientationToFixed();
        }

        public void ReorientDynamic()
        {
            PositionsLabMain.ReorientShapeOrientationToDynamic();
        }
    }
}
