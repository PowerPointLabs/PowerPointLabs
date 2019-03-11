using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.Service;

namespace PowerPointLabs.ELearningLab.Views
{
    /// <summary>
    /// Interaction logic for HumanVoiceLoginPage.xaml
    /// </summary>
    public partial class AzureVoiceLoginPage : Page
    {
        public AudioSettingsPage previousPage = AudioSettingsPage.MainSettingsPage;
        private static AzureVoiceLoginPage instance;       
        private AzureVoiceLoginPage()
        {
            InitializeComponent();
        }

        public static AzureVoiceLoginPage GetInstance()
        {
            if (instance == null)
            {
                instance = new AzureVoiceLoginPage();
            }
            return instance;
        }
        public void Destroy()
        {
            instance = null;
        }

        #region XAML-Binded Event Handlers
        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            SwitchViewToPreviousPage();
        }

        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            string _endpoint = "";
            string _key = "";

            try
            {
                _key = key.Text.Trim();
                string region = ((ComboBoxItem)endpoint.SelectedItem).Content.ToString().Trim();
                _endpoint = AzureEndpointToUriConverter.regionToEndpointMapping[region];
            }
            catch
            {
                MessageBox.Show("Key or Region cannot be empty!", "Invalid Input");
                return;
            }

            if (AzureRuntimeService.IsValidUserAccount(_key, _endpoint))
            {
                // Delete previous user account
                AzureAccount.GetInstance().Clear();
                AzureAccountStorageService.DeleteUserAccount();
                // Create and save new user account
                string _region = ((ComboBoxItem)endpoint.SelectedItem).Content.ToString().Trim();
                AzureAccount.GetInstance().SetUserKeyAndRegion(_key, _region);
                AzureAccountStorageService.SaveUserAccount(AzureAccount.GetInstance());
                AzureRuntimeService.IsAzureAccountPresentAndValid = true;
                SwitchViewToPreviousPage();
            }
            else
            {
                MessageBox.Show("Invalid Account!");
            }
        }
        #endregion

        #region Private Helper Functions

        private void SwitchViewToPreviousPage()
        {
            AudioSettingsDialogWindow dialog = AudioSettingsDialogWindow.GetInstance();
            dialog.SetCurrentPage(previousPage);
        }

        #endregion
    }
}
