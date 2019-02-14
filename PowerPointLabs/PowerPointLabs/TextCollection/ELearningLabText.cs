﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointLabs.TextCollection
{
    internal static class ELearningLabText
    {
        public const string RibbonMenuId = "eLearningLabMenu";
        public const string RibbonMenuLabel = "e-Learning";
        public const string RibbonMenuSupertip =
            "Use eLearning Lab to create audio, callouts, captions and tooltips all in one go!";

        public const string ELearningTaskPaneTag = "E-Learning Workspace";
        public const string ELearningTaskPaneSuperTip =
            "This is the work space for creating e-learning slides.";
        public const string ELearningTaskPaneLabel = "eLearningLab";

        public const string ELearningLabSettingsTag = "ELearningLabSettings";

        public const string Identifier = "PPTL";
        public const string Underscore = "_";
        public const string CalloutIdentifier = "Callout";
        public const string CaptionIdentifier = "Caption";
        public const string AudioIdentifier = "Audio";
        public const string DefaultAudioIdentifier = "Default";

        public const string AudioDefaultLabelFormat = "{0}" + Underscore + DefaultAudioIdentifier;

        public const string AudioFileNameFormat = "Slide {0} ClickNo {1} Speech.wav";

        public const string CaptionShapeNameFormat = Identifier + Underscore + "{0}" + Underscore + CaptionIdentifier;
        public const string CalloutShapeNameFormat = Identifier + Underscore + "{0}" + Underscore + CalloutIdentifier;
        public const string AudioCustomShapeNameFormat = Identifier + Underscore + "{0}" + Underscore + AudioIdentifier + Underscore + "{1}";
        public const string AudioDefaultShapeNameFormat = Identifier + Underscore + "{0}" + 
            Underscore + AudioIdentifier + Underscore + "{1}" + Underscore + DefaultAudioIdentifier;

        public const string ExtractTagNoRegex = Identifier + Underscore + @"([1-9][0-9]*)" + 
            Underscore + "(" + CalloutIdentifier + "|" + CaptionIdentifier + "|" + AudioIdentifier + @").*";
        public const string ExtractFunctionRegex = Identifier + Underscore + @"[1-9][0-9]*" + Underscore
               + "(" + CalloutIdentifier + "|" + CaptionIdentifier + "|" + AudioIdentifier + ")" + @".*";
        public const string ExtractVoiceNameRegex = Identifier + Underscore + @"[1-9][0-9]*" + Underscore
              + CalloutIdentifier + "|" + CaptionIdentifier + "|" + AudioIdentifier + Underscore + @"(.*)";
        public const string VoiceShapeNameRegex = Identifier + Underscore + @"[1-9][0-9]*" + Underscore
            + AudioIdentifier + Underscore + @".*";
        public const string CalloutShapeNameRegex = Identifier + Underscore + @"[1-9][0-9]*" + Underscore
            + CalloutIdentifier;
    }
}
