using System;
using System.Speech.Synthesis;
using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.ELearningLab.Views;

namespace PowerPointLabs.ELearningLab.AudioGenerator
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

            using (SpeechSynthesizer synthesizer = new SpeechSynthesizer())
            {
                synthesizer.SetOutputToWaveFile(directory);
                synthesizer.Speak(p);
            }
        }

        public static void Speak(PromptBuilder p)
        {
            SpeechSynthesizer synthesizer = CreateSynthesizerOutputToAudio();

            Prompt spokenPrompt = synthesizer.SpeakAsync(p);
            SynthesisState state = new SynthesisState(synthesizer, spokenPrompt);
            
            ShowSpeechCancelDialog(state);
        }

        private static SpeechSynthesizer CreateSynthesizerOutputToAudio()
        {
            SpeechSynthesizer synthesizer = new SpeechSynthesizer();
            synthesizer.SetOutputToDefaultAudioDevice();
            return synthesizer;
        }

        private static void ShowSpeechCancelDialog(SynthesisState state)
        {
            SpeechSynthesizer synthesizer = state.Synthesizer;
            Prompt spokenPrompt = state.PromptBeingSynthesized;

            SpeechPlayingDialogBox speechPlayingDialog = new SpeechPlayingDialogBox(state);
            speechPlayingDialog.Closed += (sender, e) => SpeechPlayingDialog_Closed(synthesizer, spokenPrompt);
            speechPlayingDialog.ShowThematicDialog();
        }

        private static void SpeechPlayingDialog_Closed(SpeechSynthesizer synthesizer, Prompt spokenPrompt)
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
