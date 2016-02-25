using System.Collections.Generic;

namespace PowerPointLabs.TagMatchers
{
    public class Matchers
    {
        public static IEnumerable<ITagMatcher> All
        {
            get
            {
                return new List<ITagMatcher>
                {
                    new NoEffectTagMatcher(),
                    new PauseTagMatcher(),
                    new StartVoiceTagMatcher(),
                    new EndVoiceTagMatcher(),
                    new StartSpeedTagMatcher(),
                    new EndSpeedTagMatcher(),
                    new PronounceTagMatcher()
                };
            }
        }
    }
}
