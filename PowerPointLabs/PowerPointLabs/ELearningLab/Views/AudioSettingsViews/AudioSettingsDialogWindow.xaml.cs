using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;

using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.Views.AudioSettingsViews;

namespace PowerPointLabs.ELearningLab.Views
{
    /// <summary>
    /// Interaction logic for NarrationsLabSettingsDialogBox.xaml
    /// </summary>
    public partial class AudioSettingsDialogWindow : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged = (sender, e) => { };

        public AudioSettingsWindowDisplayOptions WindowDisplayOption
        {
            get
            {
                return windowDisplayOption;
            }
            set
            {
                windowDisplayOption = value;
                PropertyChanged(this, new PropertyChangedEventArgs("WindowDisplayOption"));
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

        public Page SubAzureLoginPage
        {
            get
            {
                return subAzureLoginPage;
            }
            set
            {
                subAzureLoginPage = value;
                PropertyChanged(this, new PropertyChangedEventArgs("SubAzureLoginPage"));
            }
        }

        public Page SubWatsonLoginPage
        {
            get
            {
                return subWatsonLoginPage;
            }
            set
            {
                subWatsonLoginPage = value;
                PropertyChanged(this, new PropertyChangedEventArgs("SubWatsonLoginPage"));
            }
        }

        private Page mainPage, subAzureLoginPage, subWatsonLoginPage;
        private AudioSettingsWindowDisplayOptions windowDisplayOption;

        public AudioSettingsDialogWindow(AudioSettingsPage page)
        {
            InitializeComponent();
            mainPage = CreatePageFromIndex(page);
            mainPage.DataContext = this;
            subAzureLoginPage = new AzureVoiceLoginPage();
            subAzureLoginPage.DataContext = this;
            subWatsonLoginPage = new WatsonVoiceLoginPage();
            subWatsonLoginPage.DataContext = this;
            windowDisplayOption = AudioSettingsWindowDisplayOptions.GoToMainPage;
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

    public enum AudioSettingsWindowDisplayOptions
    {
        GoToMainPage,
        GoToAzureLoginPage,
        GoToWatsonLoginPage
    }

}
