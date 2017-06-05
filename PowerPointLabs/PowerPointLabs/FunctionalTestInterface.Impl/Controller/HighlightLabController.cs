using System;
using System.Drawing;
using System.Windows.Forms;
using PowerPointLabs.ActionFramework.Common.Extension;
using TestInterface;

namespace PowerPointLabs.FunctionalTestInterface.Impl.Controller
{
    [Serializable]
    class HighlightLabController : MarshalByRefObject, IHighlightLabController
    {
        private static IHighlightLabController _instance = new HighlightLabController();

        public static IHighlightLabController Instance { get { return _instance; } }
        
        private HighlightLabController() {}

        public void RemoveHighlighting()
        {
            UIThreadExecutor.Execute(() =>
            {
                FunctionalTestExtensions.GetRibbonUi().OnAction(
                    new RibbonControl("RemoveHighlightButton"));
            });
        }
    }
}
