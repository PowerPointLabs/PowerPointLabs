using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.TextCollection;


namespace PowerPointLabs.ELearningLab.Views
{
    /// <summary>
    /// Interaction logic for AudioSettingsPage.xaml
    /// </summary>
    public partial class AudioMainSettingsPage : Page
    {
        public delegate void DialogConfirmedDelegate(VoiceType selectedVoiceType, IVoice selectedVoice, bool isPreviewCurrentSlide);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        public delegate void DefaultVoiceChangedDelegate();
        public DefaultVoiceChangedDelegate DefaultVoiceChangedHandler { get; set; }
        public bool IsDefaultVoiceChangedHandlerAssigned { get; set; } = false;
        public VoiceType SelectedVoiceType
        {
            get
            {
                if ((bool)RadioDefaultVoice.IsChecked)
                {
                    return VoiceType.ComputerVoice;
                }
                else
                {
                    return VoiceType.AzureVoice;
                }
            }
        }

        public IVoice SelectedVoice
        {
            get
            {
                if ((bool)RadioDefaultVoice.IsChecked)
                {
                    return computerVoiceComboBox.SelectedItem as ComputerVoice;
                }
                else
                {
                    return azureVoiceComboBox.SelectedItem as AzureVoice;
                }
            }
        }

        private static AudioMainSettingsPage instance;

        private AudioMainSettingsPage()
        {
            InitializeComponent();
            azureVoiceComboBox.ItemsSource = AzureVoiceList.voices;
            azureVoiceComboBox.DisplayMemberPath = "Voice";
            computerVoiceComboBox.ItemsSource = ComputerVoiceRuntimeService.Voices;
            computerVoiceComboBox.DisplayMemberPath = "Voice";
        }
        public static AudioMainSettingsPage GetInstance()
        {
            if (instance == null)
            {
                instance = new AudioMainSettingsPage();
            }

            return instance;
        }

        public void SetAudioMainSettings(VoiceType selectedVoiceType, IVoice selectedVoice, bool isPreviewChecked)
        {
            computerVoiceComboBox.ToolTip = NarrationsLabText.SettingsVoiceSelectionInputTooltip;
            previewCheckbox.IsChecked = isPreviewChecked;
            previewCheckbox.ToolTip = NarrationsLabText.SettingsPreviewCheckboxTooltip;

            switch (selectedVoiceType)
            {
                case VoiceType.AzureVoice:
                    RadioAzureVoice.IsChecked = true;
                    azureVoiceComboBox.SelectedItem = selectedVoice as AzureVoice;
                    break;
                case VoiceType.ComputerVoice:
                    RadioDefaultVoice.IsChecked = true;
                    computerVoiceComboBox.SelectedItem = selectedVoice as ComputerVoice;
                    break;
                default:
                    break;
            }

        }

        public void Destroy()
        {
            instance = null;
        }

        #region XAML-Binded Event Handlers

        private void AudioMainSettingsPage_Loaded(object sender, RoutedEventArgs e)
        {
            RadioAzureVoice.Checked += RadioAzureVoice_Checked;
            ToggleAzureFunctionVisibility();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogConfirmedHandler(SelectedVoiceType, SelectedVoice, previewCheckbox.IsChecked.GetValueOrDefault());
            // TODO
            if (IsDefaultVoiceChangedHandlerAssigned)
            { 
                DefaultVoiceChangedHandler();
            }
            AudioSettingsDialogWindow.GetInstance().Close();
            AudioSettingsDialogWindow.GetInstance().Destroy();
        }

        private void AzureVoiceBtn_Click(object sender, RoutedEventArgs e)
        {
            AzureVoiceLoginPage.GetInstance().previousPage = AudioSettingsPage.MainSettingsPage;
            AudioSettingsDialogWindow.GetInstance().SetCurrentPage(AudioSettingsPage.AzureLoginPage);
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            if (AudioSettingService.selectedVoiceType == VoiceType.AzureVoice 
                && !AzureRuntimeService.IsAzureAccountPresent())
            {
                ComputerVoice defaultVoiceSelected = computerVoiceComboBox.SelectedItem as ComputerVoice;
                DialogConfirmedHandler(VoiceType.ComputerVoice, defaultVoiceSelected, previewCheckbox.IsChecked.GetValueOrDefault());
            }
            AudioSettingsDialogWindow.GetInstance().Destroy();
        }

        private void ChangeAccountButton_Click(object sender, RoutedEventArgs e)
        {
            AudioSettingsDialogWindow.GetInstance().SetCurrentPage(AudioSettingsPage.AzureLoginPage);
        }

        private void LogOutButton_Click(object sender, RoutedEventArgs e)
        {
            AzureAccount.GetInstance().Clear();
            AzureAccountStorageService.DeleteUserAccount();
            azureVoiceComboBox.Visibility = Visibility.Collapsed;
            azureVoiceBtn.Visibility = Visibility.Visible;
            changeAcctBtn.Visibility = Visibility.Hidden;
            logoutBtn.Visibility = Visibility.Hidden;
            RadioAzureVoice.IsEnabled = false;
            RadioDefaultVoice.IsChecked = true;
        }

        private void RadioAzureVoice_Checked(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Note that we only support English language at this stage.");
        }

        private void PreviewCheckBox_Checked(object sender, RoutedEventArgs e)
        {
            AudioSettingService.IsPreviewEnabled = true;
        }

        private void PreviewCheckBox_Unchecked(object sender, RoutedEventArgs e)
        {
            AudioSettingService.IsPreviewEnabled = false;
        }

        #endregion

        #region Private Helper Functions

        private void ToggleAzureFunctionVisibility()
        {
            if (AzureRuntimeService.IsAzureAccountPresent() && AzureRuntimeService.IsValidUserAccount())
            {
                azureVoiceComboBox.Visibility = Visibility.Visible;
                azureVoiceBtn.Visibility = Visibility.Collapsed;
                changeAcctBtn.Visibility = Visibility.Visible;
                logoutBtn.Visibility = Visibility.Visible;
                RadioAzureVoice.IsEnabled = true;
            }
            else
            {
                azureVoiceComboBox.Visibility = Visibility.Collapsed;
                azureVoiceBtn.Visibility = Visibility.Visible;
                changeAcctBtn.Visibility = Visibility.Hidden;
                logoutBtn.Visibility = Visibility.Hidden;
                RadioAzureVoice.IsEnabled = false;
            }
        }

        #endregion
    }
}
