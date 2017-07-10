using PowerPointLabs.ActionFramework.Common.Attribute;

namespace PowerPointLabs.ActionFramework.ShapesLab
{
    [ExportActionRibbonId(TextCollection.ShapesLabTag)]
    class ShapesMenuActionHandler : ShapesLabActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            InitCustomShapePane();
            TogglePaneVisibility();
        }
    }
}
