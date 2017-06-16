using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action
{
    [ExportActionRibbonId(
        TextCollection.ConvertToPictureMenuId + TextCollection.MenuShape,
        TextCollection.ConvertToPictureMenuId + TextCollection.MenuLine,
        TextCollection.ConvertToPictureMenuId + TextCollection.MenuFreeform,
        TextCollection.ConvertToPictureMenuId + TextCollection.MenuGroup,
        TextCollection.ConvertToPictureMenuId + TextCollection.MenuInk,
        TextCollection.ConvertToPictureMenuId + TextCollection.MenuVideo,
        TextCollection.ConvertToPictureMenuId + TextCollection.MenuTextEdit,
        TextCollection.ConvertToPictureMenuId + TextCollection.MenuChart,
        TextCollection.ConvertToPictureMenuId + TextCollection.MenuTable,
        TextCollection.ConvertToPictureMenuId + TextCollection.MenuTableCell,
        TextCollection.ConvertToPictureMenuId + TextCollection.MenuSmartArt,
        TextCollection.ConvertToPictureMenuId + TextCollection.MenuEditSmartArt,
        TextCollection.ConvertToPictureMenuId + TextCollection.MenuEditSmartArtText)]
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
