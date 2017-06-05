using System;
using System.Speech.Synthesis;
using System.Windows.Forms;
using PowerPointLabs.Views;

namespace PowerPointLabs.SpeechEngine
{
    class PromptToAudio
    {
        public static void SaveAsWav(PromptBuilder p, String directory)
        {
            bool hasFilePath = !String.IsNullOrWhiteSpace(directory);
            if (!hasFilePath)
            {
                // We check if there is text first, as
                // .SetOutputToWaveFile creates an empty WAV file
                // (even if nothing will be added to it.)
                return;
            }

            using (var synthesizer = new SpeechSynthesizer())
            {
                synthesizer.SetOutputToWaveFile(directory);
                synthesizer.Speak(p);
            }
        }

        public static void Speak(PromptBuilder p)
        {
            var synthesizer = CreateSynthesizerOutputToAudio();

            Prompt spokenPrompt = synthesizer.SpeakAsync(p);
            SynthesisState state = new SynthesisState(synthesizer, spokenPrompt);
            
            ShowSpeechCancelDialog(state);
        }

        private static SpeechSynthesizer CreateSynthesizerOutputToAudio()
        {
            var synthesizer = new SpeechSynthesizer();
            synthesizer.SetOutputToDefaultAudioDevice();
            return synthesizer;
        }

        private static void ShowSpeechCancelDialog(SynthesisState state)
        {
            SpeechSynthesizer synthesizer = state.Synthesizer;
            Prompt spokenPrompt = state.PromptBeingSynthesized;

            SpeechPlayingForm progress = new SpeechPlayingForm(state);
            DialogResult result = progress.ShowDialog();
            if (result == DialogResult.Cancel)
            {
                try
                {
                    synthesizer.SpeakAsyncCancel(spokenPrompt);
                }
                catch (ObjectDisposedException)
                {
                    // Synthesizer has already finished, so we don't need to do anything.
                }
            }
        }
    }
}
