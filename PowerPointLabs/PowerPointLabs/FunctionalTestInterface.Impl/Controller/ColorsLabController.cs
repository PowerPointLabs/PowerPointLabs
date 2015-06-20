using System;
using System.Drawing;
using System.Windows.Forms;
using FunctionalTestInterface;

namespace PowerPointLabs.FunctionalTestInterface.Impl.Controller
{
    [Serializable]
    class ColorsLabController : MarshalByRefObject, IColorsLabController
    {
        private static IColorsLabController _instance = new ColorsLabController();

        public static IColorsLabController Instance { get { return _instance; } }

        private ColorPane _pane;

        private ColorsLabController() {}

        public void OpenPane()
        {
            UIThreadExecutor.Execute(() =>
            {
                Globals.ThisAddIn.Ribbon.ColorPickerButtonClick(
                    new RibbonControl("ColorsLab"));
                _pane = Globals.ThisAddIn.GetActivePane(
                    typeof(ColorPane)).Control as ColorPane;
            });
        }

        public void ClosePane()
        {
            if (_pane != null)
            {
                _pane.Visible = false;
            }
        }

        public Panel GetDropletPanel()
        {
            if (_pane != null)
            {
                return _pane.GetDropletPanel();
            }
            return null;
        }

        public Button GetFontColorButton()
        {
            if (_pane != null)
            {
                return _pane.GetFontColorButton();
            }
            return null;
        }

        public Button GetLineColorButton()
        {
            if (_pane != null)
            {
                return _pane.GetLineColorButton();
            }
            return null;
        }

        public Button GetFillCollorButton()
        {
            if (_pane != null)
            {
                return _pane.GetFillCollorButton();
            }
            return null;
        }

        public Panel GetMonoPanel1()
        {
            if (_pane != null)
            {
                return _pane.GetMonoPanel1();
            }
            return null;
        }

        public Panel GetMonoPanel7()
        {
            if (_pane != null)
            {
                return _pane.GetMonoPanel7();
            }
            return null;
        }

        public Panel GetFavColorPanel1()
        {
            if (_pane != null)
            {
                return _pane.GetFavColorPanel1();
            }
            return null;
        }

        public Button GetResetFavColorsButton()
        {
            if (_pane != null)
            {
                return _pane.GetResetFavColorsButton();
            }
            return null;
        }

        public Button GetEmptyFavColorsButton()
        {
            if (_pane != null)
            {
                return _pane.GetEmptyFavColorsButton();
            }
            return null;
        }

        public IColorsLabMoreInfoDialog ShowMoreColorInfo(Color color)
        {
            if (_pane != null)
            {
                IColorsLabMoreInfoDialog dialog = null;
                UIThreadExecutor.Execute(() =>
                {
                    dialog = _pane.ShowMoreColorInfo(color);
                });
                return dialog;
            }
            return null;
        }
    }
}
