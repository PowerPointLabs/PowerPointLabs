using System;
using System.Speech.Synthesis;
using System.Windows.Forms;
using AudioGen.SpeechEngine;

namespace AudioGen.Views
{
    public partial class SpeechPlayingForm : Form
    {
        private delegate void CloseDelegate();
        public SpeechPlayingForm(SynthesisState state)
        {
            InitializeComponent();
            state.Synthesizer.SpeakCompleted += SynthesizerOnSpeakCompleted;
        }

        private void SynthesizerOnSpeakCompleted(object sender, SpeakCompletedEventArgs speakCompletedEventArgs)
        {
            if (InvokeRequired)
            {
                Invoke(new CloseDelegate(Close));
            }
            else
            {
                Close();
            }
        }

        private void CancelButtonClick(object sender, EventArgs e)
        {
            Close();
        }
    }
}
