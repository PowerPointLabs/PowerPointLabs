using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ShapesLab;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ActionFramework.ShapesLab
{
    [ExportActionRibbonId(ShapesLabText.PaneTag)]
    class ShapesLabActionHandler : ActionHandler
    {
        public CustomShapePane InitCustomShapePane()
        {
            ThisAddIn addIn = this.GetAddIn();
            addIn.InitializeShapesLabConfig();
            addIn.InitializeShapeGallery();
            addIn.RegisterShapesLabPane(this.GetCurrentPresentation().Presentation);

            var customShapePane = GetShapesLabPane();

            if (customShapePane == null)
            {
                return null;
            }

            var customShape = customShapePane.Control as CustomShapePane;

            Logger.Log(
                "Before Visible: " +
                string.Format("Pane Width = {0}, Pane Height = {1}, Control Width = {2}, Control Height {3}",
                              customShapePane.Width, customShapePane.Height, customShape.Width, customShape.Height));

            return customShape;
        }

        public Microsoft.Office.Tools.CustomTaskPane GetShapesLabPane()
        {

            var customShapePane = this.GetAddIn().GetActivePane(typeof(CustomShapePane));

            if (customShapePane == null || !(customShapePane.Control is CustomShapePane))
            {
                return null;
            }
            return customShapePane;
        }

        public void TogglePaneVisibility()
        {
            var customShapePane = GetShapesLabPane();

            if (customShapePane == null)
            {
                return;
            }

            SetPaneVisibility(!customShapePane.Visible);
        }

        public void SetPaneVisibility(bool visibility)
        {
            var customShapePane = GetShapesLabPane();

            if (customShapePane == null)
            {
                return;
            }

            customShapePane.Visible = visibility;

            if (customShapePane.Visible)
            {
                var customShape = customShapePane.Control as CustomShapePane;
                customShape.Width = customShapePane.Width - 16;
                customShape.PaneReload();
            }
        }

        protected override void ExecuteAction(string ribbonId)
        {
            InitCustomShapePane();
            TogglePaneVisibility();
        }
    }
}
