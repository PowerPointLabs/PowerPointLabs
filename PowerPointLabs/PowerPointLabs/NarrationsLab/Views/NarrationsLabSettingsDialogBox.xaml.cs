using System.ComponentModel;

using PowerPointLabs.NarrationsLab.Data;

namespace PowerPointLabs.NarrationsLab.Views
{
    /// <summary>
    /// Interaction logic for NarrationsLabSettingsDialogBox.xaml
    /// </summary>
    public partial class NarrationsLabSettingsDialogBox : INotifyPropertyChanged
    {


        public event PropertyChangedEventHandler PropertyChanged = (sender, e) => { };
        public NarrationsLabSettingsPage CurrentPage
        {
            get
            {
                return _currentPage;
            }
            set
            {
                _currentPage = value;
                PropertyChanged(this, new PropertyChangedEventArgs("CurrentPage"));
            }
        }
        private static NarrationsLabSettingsDialogBox instance;
        private NarrationsLabSettingsPage _currentPage { get; set; } = NarrationsLabSettingsPage.MainSettingsPage;
        public void SetCurrentPage(NarrationsLabSettingsPage page)
        {
            CurrentPage = page;
        }

        public static NarrationsLabSettingsDialogBox GetInstance()
        {
            if (instance == null)
            {
                instance = new NarrationsLabSettingsDialogBox();
            }
            return instance;
        }
        public void Destroy()
        {
            AzureVoiceLoginPage.GetInstance().Destroy();
            NarrationsLabMainSettingsPage.GetInstance().Destroy();
            instance = null;
        }
        private NarrationsLabSettingsDialogBox()
        {
            InitializeComponent();
            this.DataContext = this;
        }
    }
}
