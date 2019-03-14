using System.Speech.Synthesis;

namespace PowerPointLabs.ELearningLab.AudioGenerator
{
    public class SynthesisState
    {
        public SpeechSynthesizer Synthesizer { get; set; }
        public Prompt PromptBeingSynthesized { get; set; }

        public SynthesisState(SpeechSynthesizer synthesizer, Prompt promptBeingSynthesized)
        {
            Synthesizer = synthesizer;
            PromptBeingSynthesized = promptBeingSynthesized;

            synthesizer.SpeakCompleted += SynthesizerOnSpeakCompleted;
        }

        private void SynthesizerOnSpeakCompleted(object sender, SpeakCompletedEventArgs e)
        {
            SpeechSynthesizer synthesizer = sender as SpeechSynthesizer;
            synthesizer.Dispose();
        }
    }
}
