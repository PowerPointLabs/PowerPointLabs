
namespace PowerPointLabs.TooltipsLab
{
    static class TooltipsLabConstants
    {
        public const float TriggerShapeDefaultLeft = 200;
        public const float TriggerShapeDefaultTop = 200;
        public const float TriggerShapeDefaultHeight = 25;
        public const float TriggerShapeDefaultWidth = 25;
        public const float TriggerShapeAndCalloutSpacing = 10;

        public const float CalloutShapeDefaultHeight = 100;
        public const float CalloutShapeDefaultWidth = 150;

        // Explanation for the choice of constants:
        // - 0.20833 is the horizontal percentage adjustment of the arrowhead of the callout.
        //   We position the callout with middle alignment to the trigger shape, then shift it
        //   back to the right by 20.833% of the callout's width to align the arrowhead with the trigger shape.
        // - 1.125 is the vertical percentage adjustment of the arrowhead of the callout.
        //   Same explanation as the horizontal adjustment, just that this is for the height.
        public const double CalloutArrowheadHorizontalAdjustment = 0.20833;
        public const double CalloutArrowheadVerticalAdjustment = 1.125;

        public const string AnimationPaneName = "AnimationCustom";


    }
}
