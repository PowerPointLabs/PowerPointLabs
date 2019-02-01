using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.NarrationsLab.Data;
using PowerPointLabs.NarrationsLab.Storage;
using PowerPointLabs.NarrationsLab.ViewModel;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.NarrationsLab.Views
{
    /// <summary>
    /// Interaction logic for NarrationsLabMainSettingsPage.xaml
    /// </summary>
    public partial class NarrationsLabMainSettingsPage : Page
    {
        public delegate void DialogConfirmedDelegate(string voiceName, AzureVoice azureVoiceName, bool isAzureVoiceSelected, bool isPreviewing);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        private static NarrationsLabMainSettingsPage instance;

        private ObservableCollection<AzureVoice> voices = AzureVoiceList.voices;

        private NarrationsLabMainSettingsPage()
        {
            InitializeComponent();
            if (UserAccount.GetInstance().IsEmpty() || !IsValidUserAccount())
            {
                voiceList.Visibility = Visibility.Collapsed;
                azureVoiceBtn.Visibility = Visibility.Visible;
                changeAcctBtn.Visibility = Visibility.Hidden;
                logoutBtn.Visibility = Visibility.Hidden;
                RadioAzureVoice.IsEnabled = false;
            }
            else
            {
                string _key = UserAccount.GetInstance().GetKey();
                string _endpoint = UserAccount.GetInstance().GetRegion();

                voiceList.Visibility = Visibility.Visible;
                azureVoiceBtn.Visibility = Visibility.Collapsed;
                changeAcctBtn.Visibility = Visibility.Visible;
                logoutBtn.Visibility = Visibility.Visible;
                RadioAzureVoice.IsEnabled = true;
            }
            voiceList.ItemsSource = voices;
            voiceList.DisplayMemberPath = "Voice";
        }
        public static NarrationsLabMainSettingsPage GetInstance()
        {
            if (instance == null)
            {
                instance = new NarrationsLabMainSettingsPage();
            }
            else
            {
                if (UserAccount.GetInstance().IsEmpty())
                {
                    instance.voiceList.Visibility = Visibility.Collapsed;
                    instance.azureVoiceBtn.Visibility = Visibility.Visible;
                    instance.changeAcctBtn.Visibility = Visibility.Hidden;
                    instance.logoutBtn.Visibility = Visibility.Hidden;
                    instance.RadioAzureVoice.IsEnabled = false;
                }
                else
                {
                    instance.voiceList.Visibility = Visibility.Visible;
                    instance.azureVoiceBtn.Visibility = Visibility.Collapsed;
                    instance.changeAcctBtn.Visibility = Visibility.Visible;
                    instance.logoutBtn.Visibility = Visibility.Visible;
                    instance.RadioAzureVoice.IsEnabled = true;
                }
            }
            return instance;
        }

        public void SetNarrationsLabMainSettings(int selectedVoiceIndex, AzureVoice humanVoice, List<string> voices, bool isAzureVoiceSelected, bool isPreviewChecked)
        {
            voiceSelectionInput.ItemsSource = voices;
            voiceSelectionInput.ToolTip = NarrationsLabText.SettingsVoiceSelectionInputTooltip;
            voiceSelectionInput.Content = voices[selectedVoiceIndex];

            if (humanVoice != null)
            {
                voiceList.SelectedItem = humanVoice;
            }

            previewCheckbox.IsChecked = isPreviewChecked;
            previewCheckbox.ToolTip = NarrationsLabText.SettingsPreviewCheckboxTooltip;

            RadioAzureVoice.IsChecked = isAzureVoiceSelected;
            RadioDefaultVoice.IsChecked = !isAzureVoiceSelected;

        }

        public void Destroy()
        {
            instance = null;
        }

        private bool IsValidUserAccount(bool showErrorMessage = true)
        {
            try
            {
                string _key = UserAccount.GetInstance().GetKey();
                string _endpoint = EndpointToUriMapping.regionToEndpointMapping[UserAccount.GetInstance().GetRegion()];
                Authentication auth = Authentication.GetInstance(_endpoint, _key);
                string accessToken = auth.GetAccessToken();
                Console.WriteLine("Token: {0}\n", accessToken);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed authentication.");
                Console.WriteLine(ex.ToString());
                Console.WriteLine(ex.Message);
                if (showErrorMessage)
                {
                    MessageBox.Show("Failed authentication");
                }
                UserAccount.GetInstance().Clear();
                NarrationsLabStorageConfig.DeleteUserAccount();
                return false;
            }
            return true;
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string defaultVoiceSelected = RadioDefaultVoice.IsChecked == true ? voiceSelectionInput.Content.ToString() : null;
                AzureVoice azureVoiceSelected = RadioAzureVoice.IsChecked == true ? (AzureVoice)voiceList.SelectedItem : null;
                if (azureVoiceSelected == null && RadioAzureVoice.IsChecked == true)
                {
                    throw new Exception("Azure voice checked but no voice selected.");
                }
                DialogConfirmedHandler(defaultVoiceSelected, azureVoiceSelected, azureVoiceSelected != null, previewCheckbox.IsChecked.GetValueOrDefault());
            }
            catch
            {
                MessageBox.Show("Voice selected cannot be empty!", "Invalid Input");
                return;
            }
            NarrationsLabSettingsDialogBox.GetInstance().Close();
            NarrationsLabSettingsDialogBox.GetInstance().Destroy();
        }

        void VoiceSelectionInput_Item_MouseUp(object sender, MouseButtonEventArgs e)
        {
            if (e.ChangedButton == MouseButton.Left && voiceSelectionInput.IsExpanded)
            {
                string value = ((TextBlock)e.Source).Text;
                voiceSelectionInput.Content = value;
            }
        }

        private void AzureVoiceBtn_Click(object sender, RoutedEventArgs e)
        {
            NarrationsLabSettingsDialogBox.GetInstance().SetCurrentPage(NarrationsLabSettingsPage.LoginPage);
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            if (NotesToAudio.IsAzureVoiceSelected && !IsValidUserAccount(false))
            {
                string defaultVoiceSelected = voiceSelectionInput.Content.ToString();
                DialogConfirmedHandler(defaultVoiceSelected, null, false, previewCheckbox.IsChecked.GetValueOrDefault());
            }
            NarrationsLabSettingsDialogBox.GetInstance().Destroy();
        }

        private void ChangeAccountButton_Click(object sender, RoutedEventArgs e)
        {
            NarrationsLabSettingsDialogBox.GetInstance().SetCurrentPage(NarrationsLabSettingsPage.LoginPage);
        }

        private void LogOutButton_Click(object sender, RoutedEventArgs e)
        {
            UserAccount.GetInstance().Clear();
            NarrationsLabStorageConfig.DeleteUserAccount();
            voiceList.Visibility = Visibility.Collapsed;
            azureVoiceBtn.Visibility = Visibility.Visible;
            changeAcctBtn.Visibility = Visibility.Hidden;
            logoutBtn.Visibility = Visibility.Hidden;
            RadioAzureVoice.IsEnabled = false;
            RadioDefaultVoice.IsChecked = true;
        }

        private void RadioDefaultVoice_Checked(object sender, RoutedEventArgs e)
        {
            RadioAzureVoice.IsChecked = false;
        }

        private void RadioAzureVoice_Checked(object sender, RoutedEventArgs e)
        {
            RadioDefaultVoice.IsChecked = false;
            MessageBox.Show("Note that we only support English language at this stage.");
        }
    }
}
