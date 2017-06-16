using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId(
        TextCollection.ConvertToPictureId + TextCollection.MenuShape,
        TextCollection.ConvertToPictureId + TextCollection.MenuLine,
        TextCollection.ConvertToPictureId + TextCollection.MenuFreeform,
        TextCollection.ConvertToPictureId + TextCollection.MenuGroup,
        TextCollection.ConvertToPictureId + TextCollection.MenuInk,
        TextCollection.ConvertToPictureId + TextCollection.MenuVideo,
        TextCollection.ConvertToPictureId + TextCollection.MenuTextEdit,
        TextCollection.ConvertToPictureId + TextCollection.MenuChart,
        TextCollection.ConvertToPictureId + TextCollection.MenuTable,
        TextCollection.ConvertToPictureId + TextCollection.MenuTableCell,
        TextCollection.ConvertToPictureId + TextCollection.MenuSmartArt,
        TextCollection.ConvertToPictureId + TextCollection.MenuEditSmartArt,
        TextCollection.ConvertToPictureId + TextCollection.MenuEditSmartArtText)]
    class ConvertToPictureActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            var selection = this.GetCurrentSelection();
            ConvertToPicture.Convert(selection);
        }
    }
}
