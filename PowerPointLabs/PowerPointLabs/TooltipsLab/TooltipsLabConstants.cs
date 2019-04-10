using System.Drawing;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.TooltipsLab
{
    static class TooltipsLabConstants
    {
        public static readonly Size DisplayImageSize = new Size(30, 30);
        public const MsoAutoShapeType TriggerShape = MsoAutoShapeType.msoShapeOval;
        public const string ShapeNameHeader = "msoShape";
        public const string AnimationNameHeader = "msoAnimEffect";
        public const string CalloutNameSubstring = "Callout";

        public static readonly MsoAnimEffect[] AnimationEffects = new MsoAnimEffect[]
        {
            MsoAnimEffect.msoAnimEffectAppear,
            MsoAnimEffect.msoAnimEffectBounce,
            MsoAnimEffect.msoAnimEffectFade,
            MsoAnimEffect.msoAnimEffectFloat,
            MsoAnimEffect.msoAnimEffectFly,
            MsoAnimEffect.msoAnimEffectGrowAndTurn,
            MsoAnimEffect.msoAnimEffectRandomBars,
            MsoAnimEffect.msoAnimEffectPlus,
            MsoAnimEffect.msoAnimEffectSplit,
            MsoAnimEffect.msoAnimEffectSwivel,
            MsoAnimEffect.msoAnimEffectWheel,
            MsoAnimEffect.msoAnimEffectWipe,
            MsoAnimEffect.msoAnimEffectZoom
        };
        public static readonly Bitmap[] AnimationImages = new Bitmap[]
        {
            Properties.Resources.Animation_Appear,
            Properties.Resources.Animation_Bounce,
            Properties.Resources.Animation_Fade,
            Properties.Resources.Animation_Float_In,
            Properties.Resources.Animation_Fly_In,
            Properties.Resources.Animation_Grow___Turn,
            Properties.Resources.Animation_Random_Bars,
            Properties.Resources.Animation_Shape,
            Properties.Resources.Animation_Split,
            Properties.Resources.Animation_Swivel,
            Properties.Resources.Animation_Wheel,
            Properties.Resources.Animation_Wipe,
            Properties.Resources.Animation_Zoom
        };

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
