using System;
using System.Collections.Generic;
using System.ComponentModel;
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
                else if ((bool)defaultVoiceRadioButton.IsChecked)
                {
                    return VoiceType.DefaultVoice;
                }
                else
                {
                    IVoice voice = rankedAudioListView.SelectedItem as IVoice;
                    if (voice is AzureVoice)
                    {
                        return VoiceType.AzureVoice;
                    }
                    else
                    {
                        return VoiceType.ComputerVoice;
                    }
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
                else if ((bool)defaultVoiceRadioButton.IsChecked)
                {
                    return AudioSettingService.selectedVoice;
                }
                else
                {
                    return rankedAudioListView.SelectedItem as IVoice;
                }
            }
        }

        private AudioPreviewPage()
        {
            InitializeComponent();
            azureVoiceComboBox.ItemsSource = AzureVoiceList.voices
                .Where(x => !AudioSettingService.preferredVoices.Any(y => y.VoiceName == x.VoiceName));
            azureVoiceComboBox.DisplayMemberPath = "Voice";
            computerVoiceComboBox.ItemsSource = ComputerVoiceRuntimeService.Voices
                .Where(x => !AudioSettingService.preferredVoices.Any(y => y.VoiceName == x.VoiceName));
            rankedAudioListView.DataContext = this;
            rankedAudioListView.ItemsSource = AudioSettingService.preferredVoices;
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
            defaultVoiceRadioButton.Checked += RadioButton_Checked;
            azureVoiceRadioButton.Checked += RadioButton_Checked;
            computerVoiceRadioButton.Checked += RadioButton_Checked;
            ICollectionView view = CollectionViewSource.GetDefaultView(rankedAudioListView.ItemsSource);
            view.Refresh();
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
            IVoice voice = ((Button)sender).CommandParameter as IVoice;
            if (voice == null)
            {
                voice = SelectedVoice;
            }
            if (voice is ComputerVoice)
            {
                ComputerVoiceRuntimeService.SpeakString(textToSpeak, voice as ComputerVoice);
            }
            else if (voice is AzureVoice)
            {
                AzureRuntimeService.SpeakString(textToSpeak, voice as AzureVoice);
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

        private void RadioButton_Checked(object sender, RoutedEventArgs e)
        {
            if (computerVoiceRadioButton.IsChecked == true
                && computerVoiceComboBox.Items.Count > 0)
            {
                previewButton.IsEnabled = true;
                return;
            }
            else if (azureVoiceRadioButton.IsChecked == true
                && azureVoiceComboBox.Items.Count > 0)
            {
                previewButton.IsEnabled = AzureRuntimeService.IsAzureAccountPresentAndValid;
                return;
            }
            else if (azureVoiceRadioButton.IsChecked == true
                && azureVoiceComboBox.Items.Count == 0)
            {
                previewButton.IsEnabled = false;
                return;
            }

            IVoice voice;
            if (defaultVoiceRadioButton.IsChecked == true)
            {
                voice = AudioSettingService.selectedVoice;
            }
            else
            {
                voice = ((RadioButton)sender).CommandParameter as IVoice;
            }

            if (voice == null)
            {
                previewButton.IsEnabled = false;
            }
            else if (voice is ComputerVoice)
            {
                previewButton.IsEnabled = true;
            }
            else
            {
                previewButton.IsEnabled = AzureRuntimeService.IsAzureAccountPresentAndValid;
            }
        }

        #endregion

        #region Private Helper Functions
        private void ToggleAzureFunctionVisibility()
        {
            if (AzureRuntimeService.IsAzureAccountPresentAndValid)
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
