using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.AudioGenerator;

namespace PowerPointLabs.ELearningLab.Views
{
    /// <summary>
    /// Interaction logic for NarrationsLabSettingsDialogBox.xaml
    /// </summary>
    public partial class AudioSettingsDialogWindow : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged = (sender, e) => { };

        public bool GoToMainPage
        {
            get
            {
                return goToMainPage;
            }
            set
            {
                goToMainPage = value;
                PropertyChanged(this, new PropertyChangedEventArgs("GoToMainPage"));
            }
        }

        public Page MainPage
        {
            get
            {
                return mainPage;
            }
            set
            {
                mainPage = value;
                PropertyChanged(this, new PropertyChangedEventArgs("MainPage"));
            }
        }

        public Page SubPage
        {
            get
            {
                return subPage;
            }
            set
            {
                subPage = value;
                PropertyChanged(this, new PropertyChangedEventArgs("SubPage"));
            }
        }

        private Page mainPage, subPage;
        private bool goToMainPage;

        public AudioSettingsDialogWindow(AudioSettingsPage page)
        {
            InitializeComponent();
            mainPage = CreatePageFromIndex(page);
            mainPage.DataContext = this;
            subPage = new AzureVoiceLoginPage();
            subPage.DataContext = this;
            goToMainPage = true;
            DataContext = this;
            audioSettingsDialogWindow.AllowsTransparency = true;
            audioSettingsDialogWindow.Opacity = 0;
        }
        protected override void OnRenderSizeChanged(SizeChangedInfo sizeInfo)
        {
            base.OnRenderSizeChanged(sizeInfo);

            //Calculate half of the offset to move the form

            if (sizeInfo.HeightChanged)
            {
                Top += (sizeInfo.PreviousSize.Height - sizeInfo.NewSize.Height) / 2;
            }

            if (sizeInfo.WidthChanged)
            {
                Left += (sizeInfo.PreviousSize.Width - sizeInfo.NewSize.Width) / 2;
            }
        }

        private void MetroWindow_ContentRendered(object sender, System.EventArgs e)
        {
            audioSettingsDialogWindow.Opacity = 1;
        }

        private Page CreatePageFromIndex(AudioSettingsPage index)
        {
            switch (index)
            {
                case AudioSettingsPage.MainSettingsPage:
                    return new AudioMainSettingsPage();
                case AudioSettingsPage.AzureLoginPage:
                    AzureVoiceLoginPage loginInstance = new AzureVoiceLoginPage();
                    loginInstance.key.Text = "";
                    loginInstance.endpoint.SelectedIndex = -1;
                    return loginInstance;
                case AudioSettingsPage.AudioPreviewPage:
                    return new AudioPreviewPage();
                default:
                    return null;
            }
        }
    }
}
