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
        public AzureVoiceLoginPage()
        {
            InitializeComponent();
        }

        #region XAML-Binded Event Handlers
        private void Page_Loaded(object sender, RoutedEventArgs e)
        {
            key.Text = string.Empty;
            endpoint.SelectedIndex = -1;
        }

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
                _endpoint = EndpointToUriConverter.azureRegionToEndpointMapping[region];
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
                MessageBox.Show("Invalid Azure Account.\nIs your account expired?\nAre you connected to WiFi?");
            }
        }
        #endregion

        #region Private Helper Functions

        private void SwitchViewToPreviousPage()
        {
            AudioSettingsDialogWindow parentWindow = Window.GetWindow(this) as AudioSettingsDialogWindow;
            parentWindow.WindowDisplayOption = AudioSettingsWindowDisplayOptions.GoToMainPage;
        }

        #endregion
    }
}
