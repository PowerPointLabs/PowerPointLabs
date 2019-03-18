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
using PowerPointLabs.ELearningLab.Utility;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ELearningLab.Views
{
    /// <summary>
    /// Interaction logic for AudioPreviewDialogWindow.xaml
    /// </summary>
    public partial class AudioPreviewPage: Page
    {
        public delegate void DialogConfirmedDelegate(string textToSpeak, VoiceType selectedVoiceType, IVoice selectedVoice);
        public DialogConfirmedDelegate PreviewDialogConfirmedHandler { get; set; }

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
                else if ((bool)watsonVoiceRadioButton.IsChecked)
                {
                    return VoiceType.WatsonVoice;
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
                    else if (voice is WatsonVoice)
                    {
                        return VoiceType.WatsonVoice;
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
                else if ((bool)watsonVoiceRadioButton.IsChecked)
                {
                    return watsonVoiceComboBox.SelectedItem as WatsonVoice;
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

        private static Dictionary<string, string> textSpokenByPerson 
            = new Dictionary<string, string>();

        public AudioPreviewPage()
        {
            InitializeComponent();
            azureVoiceComboBox.ItemsSource = AzureVoiceList.voices
                .Where(x => !AudioSettingService.preferredVoices.Any(y => y.VoiceName == x.VoiceName));
            azureVoiceComboBox.DisplayMemberPath = "Voice";
            computerVoiceComboBox.ItemsSource = ComputerVoiceRuntimeService.Voices
                .Where(x => !AudioSettingService.preferredVoices.Any(y => y.VoiceName == x.VoiceName));
            watsonVoiceComboBox.ItemsSource = WatsonRuntimeService.Voices
                .Where(x => !AudioSettingService.preferredVoices.Any(y => y.VoiceName == x.VoiceName));
            rankedAudioListView.DataContext = this;
            rankedAudioListView.ItemsSource = AudioSettingService.preferredVoices;
        }

        #region Public Functions

        public void SetAudioPreviewSettings(string textToSpeak, VoiceType selectedVoiceType, IVoice selectedVoice)
        {
            spokenText.Text = textToSpeak;
            if (selectedVoiceType == VoiceType.DefaultVoice)
            {
                defaultVoiceRadioButton.IsChecked = true;
                return;
            }
            if (rankedAudioListView.Items.Contains(selectedVoice))
            {
                rankedAudioListView.SelectedItem = selectedVoice;
                return;
            }
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
                case VoiceType.WatsonVoice:
                    watsonVoiceRadioButton.IsChecked = true;
                    watsonVoiceComboBox.SelectedItem = selectedVoice as WatsonVoice;
                    break;
                default:
                    defaultVoiceRadioButton.IsChecked = true;
                    break;
            }
        }

        #endregion

        #region XAML-Binded Event Handlers

        private void AudioPreviewPage_Loaded(object sender, RoutedEventArgs e)
        {
            ToggleAzureFunctionVisibility();
            defaultVoiceLabel.Content = AudioSettingService.selectedVoice.ToString();
            defaultVoiceRadioButton.Checked += RadioButton_Checked;
            azureVoiceRadioButton.Checked += RadioButton_Checked;
            computerVoiceRadioButton.Checked += RadioButton_Checked;
            azureVoiceRadioButton.IsEnabled = azureVoiceComboBox.Items.Count > 0 
                && AzureRuntimeService.IsAzureAccountPresentAndValid;
            computerVoiceRadioButton.IsEnabled = computerVoiceComboBox.Items.Count > 0;
            ICollectionView view = CollectionViewSource.GetDefaultView(rankedAudioListView.ItemsSource);
            view.Refresh();
        }

        private void AzureVoiceLogInButton_Click(object sender, RoutedEventArgs e)
        {
            AudioSettingService.AudioPreviewPageHeight = Height;
            AudioSettingsDialogWindow parentWindow = Window.GetWindow(this) as AudioSettingsDialogWindow;
            parentWindow.ShouldGoToMainPage = false;
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
            else if (voice is AzureVoice && !IsSameTextSpokenBySamePerson(textToSpeak, voice.VoiceName))
            {
                AzureRuntimeService.SpeakString(textToSpeak, voice as AzureVoice);
            }
            else if (voice is AzureVoice && IsSameTextSpokenBySamePerson(textToSpeak, voice.VoiceName))
            {
                string dirPath = System.IO.Path.GetTempPath() + AudioService.TempFolderName;
                string filePath = dirPath + "\\" +
                    string.Format(ELearningLabText.AudioPreviewFileNameFormat, voice.VoiceName);
                 AzureRuntimeService.PlaySavedAudioForPreview(filePath);
            }
            else if (voice is WatsonVoice && !IsSameTextSpokenBySamePerson(textToSpeak, voice.VoiceName))
            {
                 WatsonRuntimeService.Speak(textToSpeak, voice as WatsonVoice);
            }
            else if (voice is WatsonVoice && IsSameTextSpokenBySamePerson(textToSpeak, voice.VoiceName))
            {
                string dirPath = System.IO.Path.GetTempPath() + AudioService.TempFolderName;
                string filePath = dirPath + "\\" +
                    string.Format(ELearningLabText.AudioPreviewFileNameFormat, voice.VoiceName);
                AzureRuntimeService.PlaySavedAudioForPreview(filePath);
            }
        }


        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            PreviewDialogConfirmedHandler(spokenText.Text.Trim(), SelectedVoiceType, SelectedVoice);
            AudioSettingsDialogWindow parentWindow = Window.GetWindow(this) as AudioSettingsDialogWindow;
            parentWindow.Close();
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            AudioSettingsDialogWindow parentWindow = Window.GetWindow(this) as AudioSettingsDialogWindow;
            parentWindow.Close();
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
            else if (voice is WatsonVoice)
            {
                // TODO
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

        private bool IsSameTextSpokenBySamePerson(string textToSpeak, string personName)
        {
            if (textSpokenByPerson.ContainsKey(personName))
            {
                string textSpoken = textSpokenByPerson[personName];
                if (textSpoken.Trim().Equals(textToSpeak.Trim()))
                {
                    return true;
                }
                textSpokenByPerson[personName] = textToSpeak.Trim();
                return false;
            }
            else
            {
                textSpokenByPerson.Add(personName, textToSpeak);
                return false;
            }
        }
        #endregion
    }
}
