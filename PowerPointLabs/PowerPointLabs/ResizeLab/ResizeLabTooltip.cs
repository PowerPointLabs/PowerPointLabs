
namespace PowerPointLabs.ResizeLab
{
    public class ResizeLabTooltip
    {
        #region Stretch/Shrink

        // First selected reference
        public const string StretchLeftFirstRef = "Stretch or shrink selected objects to match the left edge of the first object selected";
        public const string StretchRightFirstRef = "Stretch or shrink selected objects to match the right edge of the first object selected";
        public const string StretchTopFirstRef = "Stretch or shrink selected objects to match the top edge of the first object selected";
        public const string StretchBottomFirstRef = "Stretch or shrink selected objects to match the bottom edge of the first object selected";
        // Outermost shape reference
        public const string StretchLeftOuterRef = "Stretch or shrink selected objects to match the left edge of the leftmost object selected";
        public const string StretchRightOuterRef = "Stretch or shrink selected objects to match the right edge of the rightmost object selected";
        public const string StretchTopOuterRef = "Stretch or shrink selected objects to match the top edge of the topmost object selected";
        public const string StretchBottomOuterRef = "Stretch or shrink selected objects to match the bottom edge of the bottommost object selected";
        // Settings
        public const string StretchSettingsDialog = "Opens the settings for Stretch / Shrink functions";
        public const string StretchSettingsFirstRef = "Stretch shapes to edges of first selected shape";
        public const string StretchSettingsOuterRef = "Stretch shapes to edges of left/right/top/bottom most shape";
        
        #endregion

        #region Equalize

        public const string EqualizeWidth = "Resize selected objects to match the width of the first object selected";
        public const string EqualizeHeight = "Resize selected objects to match the height of the first object selected";
        public const string EqualizeBoth = "Resize selected objects to match the height and width of the first object selected";

        #endregion

        #region Fit To Slide

        public const string FitToSlideWidth = "Expand selected objects horizontally to fit the width of the slide";
        public const string FitToSlideHeight = "Expand selected objects vertically to fit the width of the slide";
        public const string FitToSlideFill = "Expand selected objects to fit the entire slide";

        #endregion

        #region Slight Adjust
        
        public const string SlightAdjustIncreaseWidth = "Slightly increases the width of selected objects";
        public const string SlightAdjustDecreaseWidth = "Slightly decreases the width of selected objects";
        public const string SlightAdjustIncreaseHeight = "Slightly increases the height of selected objects";
        public const string SlightAdjustDecreaseHeight = "Slightly decreases the height of selected objects";
        // Settings
        public const string SlightAdjustSettingsDialog = "Opens the settings for Adjust Slightly functions";
        public const string SlightAdjustSettingsLabel = "Defines the amount to increase or decrease the height or width by";
        public const string SlightAdjustSettingsTextBox = "Enter a decimal value greater than 0";

        #endregion

        #region Match

        public const string MatchWidth = "Resizes selected object's height to match their width";
        public const string MatchHeight = "Resizes selected object's width to match their height";

        #endregion

        #region Adjust Proportionally

        public const string AdjustProportionallyWidth = "Resizes selected object's width proportionally to the width of the first object selected";
        public const string AdjustProportionallyHeight = "Resizes selected object's height proportionally to the height of the first object selected";
        public const string AdjustProportionallyArea = "Resizes selected object's area proportionally to the area of the first object selected";
        // Settings
        public const string AdjustProportionallySettingsTextBox = "Enter a decimal value greater than 0";

        #endregion

        #region Main Settings

        public const string SettingsMaintainAspectRatio = "Maintains the aspect ratio of objects when performing resizing of objects";
        public const string SettingsAnchorPointLabel = "Lock objects at the selected point when resizing objects";
        public const string SettingsAnchorTopLeftBtn = "Top Left";
        public const string SettingsAnchorTopBtn = "Top";
        public const string SettingsAnchorTopRightBtn = "Top Right";
        public const string SettingsAnchorLeftBtn = "Left";
        public const string SettingsAnchorCenterBtn = "Center";
        public const string SettingsAnchorRightBtn = "Right";
        public const string SettingsAnchorBottomLeftBtn = "Bottom Left";
        public const string SettingsAnchorBottomBtn = "Bottom";
        public const string SettingsAnchorBottomRightBtn = "Bottom Right";

        #endregion
    }
}
