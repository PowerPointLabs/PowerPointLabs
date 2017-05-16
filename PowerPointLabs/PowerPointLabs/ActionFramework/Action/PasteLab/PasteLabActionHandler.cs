using System.Windows;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;

namespace PowerPointLabs.ActionFramework.Action.PasteLab
{
    abstract class PasteLabActionHandler : ActionHandler
    {
        // Sealed method: Subclasses should override ExecutePasteAction instead
        protected sealed override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            // Limitation: Clipboard Shapes' positions will not be preserved. Unable to find a good fix.
            IDataObject clipboardData = Clipboard.GetDataObject();

            ExecutePasteAction(ribbonId);

            if (clipboardData != null)
            {
                Clipboard.SetDataObject(clipboardData);
            }
        }

        protected abstract void ExecutePasteAction(string ribbonId);
    }
}
