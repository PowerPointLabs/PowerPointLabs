﻿using System;
using System.Speech.Synthesis;
using System.Windows;

using PowerPointLabs.SpeechEngine;

namespace PowerPointLabs.Views
{
    /// <summary>
    /// Interaction logic for SpeechPlayingDialogBox.xaml
    /// </summary>
    public partial class SpeechPlayingDialogBox
    {
        public SpeechPlayingDialogBox()
        {
            InitializeComponent();
        }
        
        public SpeechPlayingDialogBox(SynthesisState state)
            : this()
        {
            state.Synthesizer.SpeakCompleted += SynthesizerOnSpeakCompleted;
        }

        private void SynthesizerOnSpeakCompleted(object sender, SpeakCompletedEventArgs speakCompletedEventArgs)
        {
            if (!CheckAccess())
            {
                // On a different thread
                Dispatcher.Invoke(new Action(() =>
                {
                    Close();
                }));
            }
            else
            {
                Close();
            }
        }

        private void CancelButton_Click(object sender, RoutedEventArgs e)
        {
            Close();
        }
    }
}
