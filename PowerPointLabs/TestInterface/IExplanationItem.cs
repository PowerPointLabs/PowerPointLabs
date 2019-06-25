﻿namespace TestInterface
{
    public interface IExplanationItem
    {
        bool IsCallout { get; set; }
        bool IsCaption { get; set; }
        bool IsVoice { get; set; }
        string VoiceLabel { get; set; }
        bool HasShortVersion { get; set; }
        string CaptionText { get; set; }
    }
}
