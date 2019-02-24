using System;
using System.Diagnostics;
using System.Windows;
using System.Windows.Controls;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.NarrationsLab.Data;
using PowerPointLabs.NarrationsLab.Storage;
using PowerPointLabs.NarrationsLab.ViewModel;

namespace PowerPointLabs.NarrationsLab.Views
{
    /// <summary>
    /// Interaction logic for HumanVoiceLoginPage.xaml
    /// </summary>
    public partial class AzureVoiceLoginPage : Page
    {
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

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            NarrationsLabSettingsDialogBox.GetInstance()
                .SetCurrentPage(Data.NarrationsLabSettingsPage.MainSettingsPage); 
        }

        private void ConfirmButton_Click(object sender, RoutedEventArgs e)
        {
            string _endpoint = "";
            string _key = "";

            try
            {
                _key = key.Text.Trim();
                string region = ((ComboBoxItem)endpoint.SelectedItem).Content.ToString().Trim();
                _endpoint = EndpointToUriMapping.regionToEndpointMapping[region];
            }
            catch
            {
                MessageBox.Show("Key or Region cannot be empty!", "Invalid Input");
                return;
            }

            try
            {
                Authentication auth = Authentication.GetInstance(_endpoint, _key);
                string accessToken = auth.GetAccessToken();
                Console.WriteLine("Token: {0}\n", accessToken);
                Logger.Log("auth passed" + _key + " " + _endpoint);
            }
            catch (Exception ex)
            {
                Console.WriteLine("Failed authentication.");
                Console.WriteLine(ex.ToString());
                Console.WriteLine(ex.Message);
                MessageBox.Show("Failed authentication");
                Logger.Log("auth failed");
                return;
            }
            string _region = ((ComboBoxItem)endpoint.SelectedItem).Content.ToString().Trim();
            UserAccount.GetInstance().SetUserKeyAndRegion(_key, _region);
            NarrationsLabStorageConfig.SaveUserAccount(UserAccount.GetInstance());
            NarrationsLabSettingsDialogBox.GetInstance()
                .SetCurrentPage(NarrationsLabSettingsPage.MainSettingsPage);
        }
    }
}
