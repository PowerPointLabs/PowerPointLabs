using System.Drawing;

namespace PowerPointLabs.SyncLab.Views
{
    class SyncFormatPaneItemStub : SyncFormatPaneItem
    {
        public SyncFormatPaneItemStub(FormatTreeNode[] formats) :
            base(null, null, null, formats)
        {
            
        }

        new public string FormatShapeKey = null;
        new public Bitmap Image = null;
        new public bool FormatShapeExists = false;

    }
}
