
namespace PowerPointLabs.TimerLab
{
    static class TimerLabConstants
    {
        public const double DefaultDisplayDuration = 1.00;
        public const int DefaultDenomination = 10;

        public const bool DefaultCountdownSetting = false;

        public const int SizeStringLimit = 3;

        public const float DefaultTimerWidth = 600;
        public const float MinTimerWidth = 100;
        public const float MaxTimerWidth = 1000;

        public const float DefaultTimerHeight = 50;
        public const float MinTimerHeight = 10;
        public const float MaxTimerHeight = 600;

        public const float DefaultMinutesLineMarkerWidth = 4;
        public const float DefaultSecondsLineMarkerWidth = 2;

        public const float DefaultTimeMarkerHeight = 30;
        public const float DefaultTimeMarkerWidth = 1;

        public const float DefaultSliderBodyWidth = 4;
        public const float DefaultSliderHeadSize = 20;

        public const int TransparencyTransparent = 1;
        public const int Rotate180Degrees = 180;
        public const float ColorChangeDuration = 0.001f;

        public const int SecondsInMinute = 60;
        public const int StartTime = 0;

        public const double FractionalDecrementOffset = 0.60;
        public const double FractionalDecrementLowerBound = 0.00;
        public const double FractionalIncrementOffset = 0.99;
        public const double FractionalIncrementUpperBound = 0.59;

        public const string ProgressBarId = "ProgressBar";
        public const string ShapeId = "TimerLabShapeId";
        public const string TimerBodyId = "TimerBody";
        public const string TimerLineMarkerId = "TimerLineMarker";
        public const string TimerLineMarkerGroupId = "TimerLineMarkerGroup";
        public const string TimerTimeMarkerId = "TimerTimeMarker";
        public const string TimerTimeMarkerGroupId = "TimerTimeMarkerGroup";
        public const string TimerSliderHeadId = "TimerSliderHead";
        public const string TimerSliderBodyId = "TimerSliderBody";

        public const string ErrorMessageOneTimerOnly = "Only one timer allowed per slide.";
    }
}
