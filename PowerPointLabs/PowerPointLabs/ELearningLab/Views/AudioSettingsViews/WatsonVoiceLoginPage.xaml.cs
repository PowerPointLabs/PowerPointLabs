using System.Windows;
using System.Windows.Controls;

using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.AudioGenerator.WatsonVoiceGenerator.Model;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.ELearningLab.Service.StorageService;

namespace PowerPointLabs.ELearningLab.Views.AudioSettingsViews
{
    /// <summary>
    /// Interaction logic for WatsonVoiceLoginPage.xaml
    /// </summary>
    public partial class WatsonVoiceLoginPage : Page
    {
        public WatsonVoiceLoginPage()
        {
            InitializeComponent();
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
                _endpoint = EndpointToUriConverter.watsonRegionToEndpointMapping[region];
            }
            catch
            {
                MessageBox.Show("Key or Region cannot be empty!", "Invalid Input");
                return;
            }

            bool isValidAccount = WatsonRuntimeService.IsValidUserAccount(_key, _endpoint);
            if (isValidAccount)
            {
                // Delete previous user account
                WatsonAccount.GetInstance().Clear();
                WatsonAccountStorageService.DeleteUserAccount();
                // Create and save new user account
                string _region = ((ComboBoxItem)endpoint.SelectedItem).Content.ToString().Trim();
                WatsonAccount.GetInstance().SetUserKeyAndRegion(_key, _region);
                WatsonAccountStorageService.SaveUserAccount(WatsonAccount.GetInstance());
                WatsonRuntimeService.IsWatsonAccountPresentAndValid = true;
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
            AudioSettingsDialogWindow parentWindow = Window.GetWindow(this) as AudioSettingsDialogWindow;
            parentWindow.WindowDisplayOption = AudioSettingsWindowDisplayOptions.GoToMainPage;
        }

        #endregion
    }
}
