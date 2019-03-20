using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
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
        
        public ObservableCollection<IVoice> Voices { get; set; }


        private Dictionary<int, ComboBox> rankToComboBoxMapping;

        private List<IVoice> rankedAudioListCache = new List<IVoice>();

        public AudioMainSettingsPage()
        {
            InitializeComponent();
            rankToComboBoxMapping = new Dictionary<int, ComboBox>();
            azureVoiceComboBox.ItemsSource = AzureVoiceList.voices;
            azureVoiceComboBox.DisplayMemberPath = "Voice";
            computerVoiceComboBox.ItemsSource = ComputerVoiceRuntimeService.Voices;
            computerVoiceComboBox.DisplayMemberPath = "Voice";
            Voices = LoadVoices();
            audioListView.DataContext = this;
            audioListView.ItemsSource = Voices;
            preferredAudioListView.DataContext = this;
            preferredAudioListView.ItemsSource = AudioSettingService.preferredVoices;
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


        #region XAML-Binded Event Handlers

        private void AudioMainSettingsPage_Loaded(object sender, RoutedEventArgs e)
        {
            RadioAzureVoice.Checked += RadioAzureVoice_Checked;
            ToggleAzureFunctionVisibility();
            SetupAudioPreferenceUI();
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            DialogConfirmedHandler(SelectedVoiceType, SelectedVoice, previewCheckbox.IsChecked.GetValueOrDefault());
            if (IsDefaultVoiceChangedHandlerAssigned)
            { 
                DefaultVoiceChangedHandler();
            }
            if (audioListView.IsVisible && !UpdateRankedAudioPreferences())
            {
                return;
            }
            AudioSettingStorageService.SaveAudioSettingPreference();
            AudioSettingsDialogWindow parentWindow = Window.GetWindow(this) as AudioSettingsDialogWindow;
            parentWindow.Close();
            SelfExplanationBlockView.dialog.NarrationsLabSettingsMainFrame.Refresh();
        }

        private void AzureVoiceBtn_Click(object sender, RoutedEventArgs e)
        {
            AudioSettingsDialogWindow parentWindow = Window.GetWindow(this) as AudioSettingsDialogWindow;
            parentWindow.ShouldGoToMainPage = false;
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            if (audioListView.IsVisible)
            {
                List<IVoice> rankedAudioListCacheCopy = rankedAudioListCache.Select(x => (IVoice)x.Clone()).ToList();
                AudioSettingService.preferredVoices = new ObservableCollection<IVoice>(rankedAudioListCacheCopy);
            }
            if (AudioSettingService.selectedVoiceType == VoiceType.AzureVoice 
                && !AzureRuntimeService.IsAzureAccountPresent())
            {
                ComputerVoice defaultVoiceSelected = computerVoiceComboBox.SelectedItem as ComputerVoice;
                DialogConfirmedHandler(VoiceType.ComputerVoice, defaultVoiceSelected, previewCheckbox.IsChecked.GetValueOrDefault());
            }
        }

        private void ChangeAccountButton_Click(object sender, RoutedEventArgs e)
        {
            AudioSettingsDialogWindow parentWindow = Window.GetWindow(this) as AudioSettingsDialogWindow;
            parentWindow.ShouldGoToMainPage = false;
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
            AzureRuntimeService.IsAzureAccountPresentAndValid = false;
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

        private void EditRankingButton_Clicked(object sender, RoutedEventArgs e)
        {
            rankedAudioListCache = AudioSettingService.preferredVoices.Select(x => (IVoice)x.Clone()).ToList();
            editPreferenceButton.Visibility = Visibility.Collapsed;
            audioListView.Visibility = Visibility.Visible;
            updatePreferenceButton.Visibility = Visibility.Visible;
            preferredAudioListView.Visibility = Visibility.Collapsed;
        }

        /// <summary>
        /// When ranking combobox selection changed, we would like to check if there exists 
        /// duplicate ranking. For example, if item A has ranking 1, and subsequently item B's
        /// ranking is also changed to ranking 1. Then we should remove ranking 1 for item A.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void RankingComboBox_SelectionChanged(object sender, RoutedEventArgs e)
        {
            ComboBox comboBox = sender as ComboBox;
            if (comboBox.SelectedIndex > 0)
            {
                int rank = comboBox.SelectedIndex;
                // check if a combobox with the same ranking already exists
                if (rankToComboBoxMapping.ContainsKey(rank))
                {
                    ComboBox previousComboBox = rankToComboBoxMapping[rank];
                    // if another combobox item with the same ranking exists, we will 
                    // remove the ranking for the previous combobox item.
                    if (!comboBox.Equals(previousComboBox))
                    {
                        previousComboBox.SelectedIndex = 0;
                    }
                    rankToComboBoxMapping.Remove(rank);
                }
                rankToComboBoxMapping.Add(rank, comboBox);
            }
        }

        private void UpdateRankingButton_Clicked(object sender, RoutedEventArgs e)
        {
            if (!UpdateRankedAudioPreferences())
            {
                return;
            }
            preferredAudioListView.ItemsSource = null;
            preferredAudioListView.ItemsSource = AudioSettingService.preferredVoices;

            // update UI
            editPreferenceButton.Visibility = Visibility.Visible;
            audioListView.Visibility = Visibility.Collapsed;
            updatePreferenceButton.Visibility = Visibility.Collapsed;
            preferredAudioListView.Visibility = Visibility.Visible;
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

        private ObservableCollection<IVoice> LoadVoices()
        {
            ObservableCollection<IVoice> voices = new ObservableCollection<IVoice>();
            foreach (IVoice voice in AzureVoiceList.voices)
            {
                voices.Add(voice);
            }
            foreach (IVoice voice in ComputerVoiceRuntimeService.Voices)
            {
                voices.Add(voice);
            }
            return voices;
        }

        private bool UpdateRankedAudioPreferences()
        {
            List<IVoice> voices = Voices.ToList().Where(x => x.Rank > 0).OrderBy(x => x.Rank).ToList();
            ObservableCollection<IVoice> voicesRanked = new ObservableCollection<IVoice>();
            for (int i = 0; i < voices.Count; i++)
            {
                IVoice voice = voices[i];
                if (voice.Rank != i + 1)
                {
                    MessageBox.Show("Please rank in sequence.");
                    return false;
                }
            }
            foreach (IVoice voice in voices)
            {
                voicesRanked.Add(voice);
            }
            AudioSettingService.preferredVoices = voicesRanked;
            return true;
        }

        private void SetupAudioPreferenceUI()
        {
            editPreferenceButton.Visibility = Visibility.Visible;
            audioListView.Visibility = Visibility.Collapsed;
            updatePreferenceButton.Visibility = Visibility.Collapsed;
            preferredAudioListView.Visibility = Visibility.Visible;
        }
        #endregion
    }
}
