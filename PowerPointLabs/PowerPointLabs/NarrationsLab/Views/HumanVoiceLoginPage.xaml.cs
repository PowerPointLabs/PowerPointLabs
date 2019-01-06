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
    public partial class HumanVoiceLoginPage : Page
    {
        private static HumanVoiceLoginPage instance;
        private HumanVoiceLoginPage()
        {
            InitializeComponent();
        }

        public static HumanVoiceLoginPage GetInstance()
        {
            if (instance == null)
            {
                instance = new HumanVoiceLoginPage();
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
            string _key = key.Text.Trim();
            string region = ((ComboBoxItem)endpoint.SelectedItem).Content.ToString().Trim();
            string _endpoint = EndpointToUriMapping.regionToEndpointMapping[region];

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
            
            UserAccount.GetInstance().SetUserKeyAndRegion(_key, region);
            NarrationsLabStorageConfig.SaveUserAccount(UserAccount.GetInstance());
            NarrationsLabSettingsDialogBox.GetInstance()
                .SetCurrentPage(NarrationsLabSettingsPage.MainSettingsPage);
        }
    }
}
