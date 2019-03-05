using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.Service;

namespace PowerPointLabs.ELearningLab.Views
{
    /// <summary>
    /// Interaction logic for AudioPreviewDialogWindow.xaml
    /// </summary>
    public partial class AudioPreviewPage: Page
    {
        public delegate void DialogConfirmedDelegate(string textToSpeak, VoiceType selectedVoiceType, IVoice selectedVoice);
        public DialogConfirmedDelegate PreviewDialogConfirmedHandler { get; set; }
        public static AudioPreviewPage GetInstance()
        {
            if (instance == null)
            {
                instance = new AudioPreviewPage();
            }

            return instance;
        }

        public VoiceType SelectedVoiceType
        {
            get
            {
                if ((bool)azureVoiceRadioButton.IsChecked)
                {
                    return VoiceType.AzureVoice;
                }
                else if ((bool)computerVoiceRadioButton.IsChecked)
                {
                    return VoiceType.ComputerVoice;
                }
                else
                {
                    return VoiceType.DefaultVoice;
                }
            }
        }

        public IVoice SelectedVoice
        {
            get
            {
                if ((bool)azureVoiceRadioButton.IsChecked)
                {
                    return azureVoiceComboBox.SelectedItem as AzureVoice;
                }
                else if ((bool)computerVoiceRadioButton.IsChecked)
                {
                    return computerVoiceComboBox.SelectedItem as ComputerVoice;
                }
                else
                {
                    return AudioSettingService.selectedVoice;
                }
            }
        }

        private AudioPreviewPage()
        {
            InitializeComponent();
            azureVoiceComboBox.ItemsSource = AzureVoiceList.voices;
            azureVoiceComboBox.DisplayMemberPath = "Voice";
            computerVoiceComboBox.ItemsSource = ComputerVoiceRuntimeService.Voices;
        }

        public void Destroy()
        {
            instance = null;
        }

        private static AudioPreviewPage instance;

        #region Public Functions

        public void SetAudioPreviewSettings(string textToSpeak, VoiceType selectedVoiceType, IVoice selectedVoice)
        {
            spokenText.Text = textToSpeak;
            switch (selectedVoiceType)
            {
                case VoiceType.AzureVoice:
                    azureVoiceRadioButton.IsChecked = true;
                    azureVoiceComboBox.SelectedItem = selectedVoice as AzureVoice;
                    break;
                case VoiceType.ComputerVoice:
                    computerVoiceRadioButton.IsChecked = true;
                    computerVoiceComboBox.SelectedItem = selectedVoice as ComputerVoice;
                    break;
                case VoiceType.DefaultVoice:
                default:
                    defaultVoiceRadioButton.IsChecked = true;
                    break;
            }
        }

        #endregion

        #region XAML-Binded Event Handlers

        private void AudioPreviewPage_Loaded(object sender, RoutedEventArgs e)
        {
            instance.ToggleAzureFunctionVisibility();
            defaultVoiceLabel.Content = AudioSettingService.selectedVoice.ToString();
        }

        private void AzureVoiceLogInButton_Click(object sender, RoutedEventArgs e)
        {
            AzureVoiceLoginPage.GetInstance().previousPage = AudioSettingsPage.AudioPreviewPage;
            AudioSettingService.AudioPreviewPageHeight = Height;
            AudioSettingsDialogWindow.GetInstance().SetDialogWindowHeight(AudioSettingService.AudioMainSettingsPageHeight);
            AudioSettingsDialogWindow.GetInstance().SetCurrentPage(AudioSettingsPage.AzureLoginPage);
        }

        private void SpeakButton_Click(object sender, RoutedEventArgs e)
        {
            string textToSpeak = spokenText.Text.Trim();
            if (string.IsNullOrEmpty(textToSpeak))
            {
                return;
            }
            if (SelectedVoice is ComputerVoice)
            {
                ComputerVoiceRuntimeService.SpeakString(textToSpeak, SelectedVoice as ComputerVoice);
            }
            else if (SelectedVoice is AzureVoice)
            {
                AzureRuntimeService.SpeakString(textToSpeak, SelectedVoice as AzureVoice);
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            PreviewDialogConfirmedHandler(spokenText.Text.Trim(), SelectedVoiceType, SelectedVoice);
            AudioSettingsDialogWindow.GetInstance().Close();
            AudioSettingsDialogWindow.GetInstance().Destroy();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            AudioSettingsDialogWindow.GetInstance().Close();
            AudioSettingsDialogWindow.GetInstance().Destroy();
        }

        #endregion

        #region Private Helper Functions
        private void ToggleAzureFunctionVisibility()
        {
            if (AzureRuntimeService.IsAzureAccountPresent() && AzureRuntimeService.IsValidUserAccount(showErrorMessage: false))
            {
                azureVoiceComboBox.Visibility = Visibility.Visible;
                azureVoiceLoginButton.Visibility = Visibility.Collapsed;
                azureVoiceRadioButton.IsEnabled = true;
            }
            else
            {
                azureVoiceComboBox.Visibility = Visibility.Collapsed;
                azureVoiceLoginButton.Visibility = Visibility.Visible;
                azureVoiceRadioButton.IsEnabled = false;
            }
        }
        #endregion
    }
}
