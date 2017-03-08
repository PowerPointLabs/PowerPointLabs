using System.Windows.Forms;

namespace PowerPointLabs.SyncLab.View
{
    public partial class SyncPane : UserControl
    {

        private bool _firstTimeLoading = true;

        public SyncPane()
        {
            InitializeComponent();
        }

        #region API
        public void PaneReload(bool forceReload = false)
        {
            if (!_firstTimeLoading && !forceReload)
            {
                return;
            }

            _firstTimeLoading = false;
        }
        #endregion

    }
}
