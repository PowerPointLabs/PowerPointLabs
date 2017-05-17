using System.Windows;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Util;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    abstract class PasteLabActionHandler : BaseUtilActionHandler
    {
        // Sealed method: Subclasses should override ExecutePasteAction instead
        protected sealed override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            // Limitation: Clipboard's shape positions will not be preserved. Unable to find a good fix.
            IDataObject clipboardData = Clipboard.GetDataObject();
            bool isClipboardEmpty = clipboardData == null || clipboardData.GetFormats().Length == 0;

            ExecutePasteAction(ribbonId, isClipboardEmpty);

            if (!isClipboardEmpty)
            {
                Clipboard.SetDataObject(clipboardData);
            }
        }

        protected abstract void ExecutePasteAction(string ribbonId, bool isClipboardEmpty);
    }
}
