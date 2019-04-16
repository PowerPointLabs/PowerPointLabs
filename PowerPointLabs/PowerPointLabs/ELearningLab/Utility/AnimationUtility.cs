using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ELearningLab.Utility
{
    public class AnimationUtility
    {
        /// <summary>
        /// Append animation for `shape` with effect type `effectType` to `slide`
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="shape"></param>
        /// <param name="effectType">specifies the type for the generated animation effect</param>
        /// <param name="clickNo">the clickNo in which the specified effect should appear</param>
        /// <returns></returns>
        public static Effect AppendAnimationToSlide(PowerPointSlide slide, Shape shape, 
            MsoAnimEffect effectType, int clickNo)
        {
            Effect effect = slide.SetShapeAsClickTriggered(shape, clickNo, effectType);
            if (StringUtility.ExtractFunctionFromString(shape.Name).Equals(ELearningLabText.CalloutIdentifier))
            {
                effect.EffectType = MsoAnimEffect.msoAnimEffectFade;
            }
            return effect;
        }
    }
}
