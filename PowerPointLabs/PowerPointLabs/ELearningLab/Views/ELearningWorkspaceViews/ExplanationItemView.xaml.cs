using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.ELearningLab.AudioGenerator;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.ELearningLab.Utility;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

namespace PowerPointLabs.ELearningLab.Views
{
    /// <summary>
    /// Interaction logic for SelfExplanationBlockView.xaml
    /// </summary>
    public partial class ExplanationItemView : UserControl
    {
        #region Custom Events

        public static readonly RoutedEvent UpButtonClickedEvent = EventManager.RegisterRoutedEvent(
            "UpButtonClickedHandler",
            RoutingStrategy.Bubble,
            typeof(RoutedEventHandler),
            typeof(ExplanationItemView));

        public static readonly RoutedEvent DownButtonClickedEvent = EventManager.RegisterRoutedEvent(
            "DownButtonClickedHandler",
            RoutingStrategy.Bubble,
            typeof(RoutedEventHandler),
            typeof(ExplanationItemView));

        public static readonly RoutedEvent DeleteButtonClickedEvent = EventManager.RegisterRoutedEvent(
           "DeleteButtonClickedHandler",
           RoutingStrategy.Bubble,
           typeof(RoutedEventHandler),
           typeof(ExplanationItemView));

        public static readonly RoutedEvent TriggerTypeSelectionChangedEvent = EventManager.RegisterRoutedEvent(
            "TriggerTypeSelectionChangedHandler",
            RoutingStrategy.Bubble,
            typeof(RoutedEventHandler),
            typeof(ExplanationItemView));

        #endregion

        #region Event Handler

        public event RoutedEventHandler UpButtonClickedHandler
        {
            add { AddHandler(UpButtonClickedEvent, value); }
            remove { RemoveHandler(UpButtonClickedEvent, value); }
        }

        public event RoutedEventHandler DownButtonClickedHandler
        {
            add { AddHandler(DownButtonClickedEvent, value); }
            remove { RemoveHandler(DownButtonClickedEvent, value); }
        }

        public event RoutedEventHandler DeleteButtonClickedHandler
        {
            add { AddHandler(DeleteButtonClickedEvent, value); }
            remove { RemoveHandler(DeleteButtonClickedEvent, value); }
        }

        public event RoutedEventHandler TriggerTypeSelectionChangedHandler
        {
            add { AddHandler(TriggerTypeSelectionChangedEvent, value); }
            remove { RemoveHandler(TriggerTypeSelectionChangedEvent, value); }
        }

        #endregion

        public ExplanationItemView()
        {
            InitializeComponent();
            upImage.Source = CommonUtil.CreateBitmapSource(Properties.Resources.Up);
            deleteImage.Source = CommonUtil.CreateBitmapSource(Properties.Resources.SyncLabDeleteButton);
            downImage.Source = CommonUtil.CreateBitmapSource(Properties.Resources.Down);
            audioImage.Source = CommonUtil.CreateBitmapSource(Properties.Resources.SpeakTextContext);
            cancelCalloutImage.Source = CommonUtil.CreateBitmapSource(Properties.Resources.CancelCalloutButton);
        }

        #region XAML-Binded Action Handler

        private void SelfExplanationBlockView_Loaded(object sender, RoutedEventArgs e)
        {
            triggerTypeComboBox.SelectionChanged += TriggerTypeComboBox_SelectionChanged;
            audioCheckBox.Checked += AudioCheckBox_CheckedChanged;
            audioCheckBox.Unchecked += AudioCheckBox_CheckedChanged;
            audioPreviewButton.IsEnabled = (bool)audioCheckBox.IsChecked;
            audioImage.Opacity = (bool)audioCheckBox.IsChecked ? 1 : 0.5;
            audioCheckBox.Unchecked += IsVoiceCaptionCalloutCheckBox_CheckChanged;
            captionCheckBox.Unchecked += IsVoiceCaptionCalloutCheckBox_CheckChanged;
            calloutCheckBox.Unchecked += IsVoiceCaptionCalloutCheckBox_CheckChanged;
            audioCheckBox.Checked += IsVoiceCaptionCalloutCheckBox_CheckChanged;
            captionCheckBox.Checked += IsVoiceCaptionCalloutCheckBox_CheckChanged;
            calloutCheckBox.Checked += IsVoiceCaptionCalloutCheckBox_CheckChanged;
        }

        private void UpButton_Click(object sender, RoutedEventArgs e)
        {
            RoutedEventArgs eventArgs = new RoutedEventArgs(UpButtonClickedEvent);
            eventArgs.Source = sender;
            RaiseEvent(eventArgs);
        }

        private void DownButton_Click(object sender, RoutedEventArgs e)
        {
            RoutedEventArgs eventArgs = new RoutedEventArgs(DownButtonClickedEvent);
            eventArgs.Source = sender;
            RaiseEvent(eventArgs);
        }

        private void DeleteButton_Click(object sender, RoutedEventArgs e)
        {
            RoutedEventArgs eventArgs = new RoutedEventArgs(DeleteButtonClickedEvent);
            eventArgs.Source = sender;
            RaiseEvent(eventArgs);
        }

        private void VoicePreviewButton_Click(object sender, RoutedEventArgs e)
        {
            AzureAccountStorageService.LoadUserAccount();
            AudioSettingsDialogWindow dialog = new AudioSettingsDialogWindow(AudioSettingsPage.AudioPreviewPage);
            AudioPreviewPage page = dialog.MainPage as AudioPreviewPage;
            page.PreviewDialogConfirmedHandler = OnSettingsDialogConfirmed;
            ConfigureAudioPreviewSettings(page);
            dialog.Title = "Audio Preview Window";
            dialog.ShowThematicDialog();
        }

        private void ShorterCalloutCancelButton_Click(object sender, RoutedEventArgs e)
        {
            calloutTextBox.Visibility = Visibility.Collapsed;
            hasShortVersionCheckBox.Visibility = Visibility.Visible;
            cancelCalloutBorder.Visibility = Visibility.Collapsed;
            hasShortVersionCheckBox.IsChecked = false;
        }

        private void TriggerTypeComboBox_SelectionChanged(object sender, RoutedEventArgs e)
        {
            RoutedEventArgs eventArgs = new RoutedEventArgs(TriggerTypeSelectionChangedEvent);
            eventArgs.Source = sender;
            RaiseEvent(eventArgs);
        }

        private void HasShortVersionCheckBox_CheckChanged(object sender, RoutedEventArgs e)
        {
            if ((bool)((CheckBox)sender).IsChecked)
            {
                calloutTextBox.Visibility = Visibility.Visible;
                hasShortVersionCheckBox.Visibility = Visibility.Collapsed;
                cancelCalloutBorder.Visibility = Visibility.Visible;
            }
            else
            {
                calloutTextBox.Visibility = Visibility.Collapsed;
                hasShortVersionCheckBox.Visibility = Visibility.Visible;
                cancelCalloutBorder.Visibility = Visibility.Collapsed;
            }
        }

        private void TriggerTypeComboBox_IsEnabledChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            ComboBox comboBox = (ComboBox)sender;
            if (!comboBox.IsEnabled)
            {
                comboBox.Opacity = 0.5;
            }
            else
            {
                comboBox.Opacity = 1;
            }
        }

        private void AudioCheckBox_CheckedChanged(object sender, RoutedEventArgs e)
        {
            if ((bool)((CheckBox)sender).IsChecked)
            {
                audioNameLabel.Visibility = Visibility.Visible;
                audioNameLabel.Text = string.Format(ELearningLabText.AudioDefaultLabelFormat, 
                    AudioSettingService.selectedVoice.ToString());
                audioPreviewButton.IsEnabled = true;
                audioImage.Opacity = 1;
            }
            else
            {
                audioNameLabel.Visibility = Visibility.Collapsed;
                audioPreviewButton.IsEnabled = false;
                audioImage.Opacity = 0.5;
            }
        }

        private void IsVoiceCaptionCalloutCheckBox_CheckChanged(object sender, RoutedEventArgs e)
        {
            RoutedEventArgs eventArgs = new RoutedEventArgs(TriggerTypeSelectionChangedEvent);
            eventArgs.Source = sender;
            RaiseEvent(eventArgs);
        }

        #endregion

        #region Private Helpder
        private void OnSettingsDialogConfirmed(string textToPreview, VoiceType selectedVoiceType, IVoice selectedVoice)
        {
            captionTextBox.Text = textToPreview;
            switch (selectedVoiceType)
            {
                case VoiceType.AzureVoice:
                case VoiceType.ComputerVoice:
                case VoiceType.WatsonVoice:
                    audioNameLabel.Text = selectedVoice.ToString();
                    break;
                case VoiceType.DefaultVoice:
                    audioNameLabel.Text = string.Format(ELearningLabText.AudioDefaultLabelFormat, 
                        AudioSettingService.selectedVoice.ToString());
                    break;
                default:
                    break;
            }
        }

        private void ConfigureAudioPreviewSettings(AudioPreviewPage page)
        {
            string textToSpeak = captionTextBox.Text.Trim();
            string voiceName = StringUtility.ExtractVoiceNameFromVoiceLabel(audioNameLabel.Text.ToString());
            if (!AudioService.CheckIfVoiceExists(voiceName))
            {
                page.SetAudioPreviewSettings(textToSpeak, VoiceType.DefaultVoice, AudioSettingService.selectedVoice);
                return;
            }
            string defaultPostfix = StringUtility.ExtractDefaultLabelFromVoiceLabel(audioNameLabel.Text.ToString());
            VoiceType voiceType = AudioService.GetVoiceTypeFromString(voiceName, defaultPostfix);
            IVoice voice = AudioService.GetVoiceFromString(voiceName);
            page.SetAudioPreviewSettings(textToSpeak, voiceType, voice);
        }

        #endregion

    }
}
