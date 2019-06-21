using System.Windows;
using System.Windows.Controls;

using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.AudioGenerator.WatsonVoiceGenerator.Model;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.ELearningLab.Service.StorageService;
using PowerPointLabs.Utils.Windows;

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
                _endpoint = EndpointToUriConverter.watsonRegionToEndpointMapping[region];
            }
            catch
            {
                MessageBoxUtil.Show("Key or Region cannot be empty!", "Invalid Input");
                return;
            }

            bool isValidAccount = WatsonRuntimeService.IsValidUserAccount(_key, _endpoint,
                "Invalid Watson Account.\nIs your Watson account expired?\nAre you connected to Wifi?");
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
                MessageBoxUtil.Show("Invalid Watson Account.\nIs your Watson account expired?\nAre you connected to Wifi?");
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
