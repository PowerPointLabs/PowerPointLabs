using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;

namespace PowerPointLabs.ELearningLab.Converters
{
    public class EffectToAnimationTypeConverter
    {
        private static HashSet<MsoAnimEffect> motionPathEffectCollection = new HashSet<MsoAnimEffect>
            {
            MsoAnimEffect.msoAnimEffectPath4PointStar, MsoAnimEffect.msoAnimEffectPath5PointStar,
            MsoAnimEffect.msoAnimEffectPath6PointStar, MsoAnimEffect.msoAnimEffectPath8PointStar,
            MsoAnimEffect.msoAnimEffectPathArcDown, MsoAnimEffect.msoAnimEffectPathArcLeft,
            MsoAnimEffect.msoAnimEffectPathArcRight, MsoAnimEffect.msoAnimEffectPathArcUp,
            MsoAnimEffect.msoAnimEffectPathBean, MsoAnimEffect.msoAnimEffectPathBounceLeft,
            MsoAnimEffect.msoAnimEffectPathBounceRight, MsoAnimEffect.msoAnimEffectPathBuzzsaw,
            MsoAnimEffect.msoAnimEffectPathCircle, MsoAnimEffect.msoAnimEffectPathCrescentMoon,
            MsoAnimEffect.msoAnimEffectPathCurvedSquare, MsoAnimEffect.msoAnimEffectPathCurvedX,
            MsoAnimEffect.msoAnimEffectPathCurvyLeft, MsoAnimEffect.msoAnimEffectPathCurvyRight,
            MsoAnimEffect.msoAnimEffectPathCurvyStar, MsoAnimEffect.msoAnimEffectPathDecayingWave,
            MsoAnimEffect.msoAnimEffectPathDiagonalDownRight, MsoAnimEffect.msoAnimEffectPathDiagonalUpRight,
            MsoAnimEffect.msoAnimEffectPathDiamond, MsoAnimEffect.msoAnimEffectPathDown,
            MsoAnimEffect.msoAnimEffectPathEqualTriangle, MsoAnimEffect.msoAnimEffectPathFigure8Four,
            MsoAnimEffect.msoAnimEffectPathFootball, MsoAnimEffect.msoAnimEffectPathFunnel,
            MsoAnimEffect.msoAnimEffectPathHeart, MsoAnimEffect.msoAnimEffectPathHeartbeat,
            MsoAnimEffect.msoAnimEffectPathHexagon, MsoAnimEffect.msoAnimEffectPathHorizontalFigure8,
            MsoAnimEffect.msoAnimEffectPathInvertedSquare, MsoAnimEffect.msoAnimEffectPathInvertedTriangle,
            MsoAnimEffect.msoAnimEffectPathLeft, MsoAnimEffect.msoAnimEffectPathLoopdeLoop,
            MsoAnimEffect.msoAnimEffectPathNeutron, MsoAnimEffect.msoAnimEffectPathOctagon,
            MsoAnimEffect.msoAnimEffectPathParallelogram, MsoAnimEffect.msoAnimEffectPathPeanut,
            MsoAnimEffect.msoAnimEffectPathPentagon, MsoAnimEffect.msoAnimEffectPathPlus,
            MsoAnimEffect.msoAnimEffectPathPointyStar, MsoAnimEffect.msoAnimEffectPathRight,
            MsoAnimEffect.msoAnimEffectPathRightTriangle, MsoAnimEffect.msoAnimEffectPathSCurve1,
            MsoAnimEffect.msoAnimEffectPathSCurve2, MsoAnimEffect.msoAnimEffectPathSineWave,
            MsoAnimEffect.msoAnimEffectPathSpiralLeft, MsoAnimEffect.msoAnimEffectPathSpiralRight,
            MsoAnimEffect.msoAnimEffectPathSpring, MsoAnimEffect.msoAnimEffectPathSquare,
            MsoAnimEffect.msoAnimEffectPathStairsDown, MsoAnimEffect.msoAnimEffectPathSwoosh,
            MsoAnimEffect.msoAnimEffectPathTeardrop, MsoAnimEffect.msoAnimEffectPathTrapezoid,
            MsoAnimEffect.msoAnimEffectPathTurnDown, MsoAnimEffect.msoAnimEffectPathTurnRight,
            MsoAnimEffect.msoAnimEffectPathTurnUp, MsoAnimEffect.msoAnimEffectPathTurnUpRight,
            MsoAnimEffect.msoAnimEffectPathUp, MsoAnimEffect.msoAnimEffectPathVerticalFigure8,
            MsoAnimEffect.msoAnimEffectPathWave, MsoAnimEffect.msoAnimEffectPathZigzag
        };
        private static HashSet<MsoAnimEffect> emphasisEffectCollection = new HashSet<MsoAnimEffect>
        {
            MsoAnimEffect.msoAnimEffectChangeFillColor, MsoAnimEffect.msoAnimEffectChangeFontColor,
            MsoAnimEffect.msoAnimEffectGrowShrink, MsoAnimEffect.msoAnimEffectVerticalGrow, 
            MsoAnimEffect.msoAnimEffectGrowWithColor, MsoAnimEffect.msoAnimEffectGrowAndTurn,
            MsoAnimEffect.msoAnimEffectBoldFlash, MsoAnimEffect.msoAnimEffectBrushOnColor,
            MsoAnimEffect.msoAnimEffectComplementaryColor, MsoAnimEffect.msoAnimEffectComplementaryColor2,
            MsoAnimEffect.msoAnimEffectContrastingColor, MsoAnimEffect.msoAnimEffectDarken,
            MsoAnimEffect.msoAnimEffectDesaturate, MsoAnimEffect.msoAnimEffectLighten,
            MsoAnimEffect.msoAnimEffectBrushOnUnderline, MsoAnimEffect.msoAnimEffectShimmer,
            MsoAnimEffect.msoAnimEffectTeeter, MsoAnimEffect.msoAnimEffectBoldReveal,
            MsoAnimEffect.msoAnimEffectWave
        };

        public static AnimationType GetAnimationTypeOfEffect(Effect effect)
        {
            if (effect.Exit == Microsoft.Office.Core.MsoTriState.msoTrue)
            {
                return AnimationType.Exit;
            }
            else if (motionPathEffectCollection.Contains(effect.EffectType))
            {
                return AnimationType.MotionPath;
            }
            else if (emphasisEffectCollection.Contains(effect.EffectType))
            {
                return AnimationType.Emphasis;
            }
            else
            {
                return AnimationType.Entrance;
            }
        }
    }
}
