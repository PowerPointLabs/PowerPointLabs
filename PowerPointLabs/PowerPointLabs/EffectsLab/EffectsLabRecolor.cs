using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;

namespace PowerPointLabs.EffectsLab
{
    internal static class EffectsLabRecolor
    {
        public static void GrayScaleRemainderEffect(PowerPointSlide curSlide, Selection selection)
        {
            PowerPointBgEffectSlide effectSlide = EffectsLabUtil.GenerateEffectSlide(curSlide, selection, true);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.GrayScaleBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public static void BlackWhiteRemainderEffect(PowerPointSlide curSlide, Selection selection)
        {
            PowerPointBgEffectSlide effectSlide = EffectsLabUtil.GenerateEffectSlide(curSlide, selection, true);
            
            if (effectSlide == null)
            {
                return;
            }

            effectSlide.BlackWhiteBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public static void GothamRemainderEffect(PowerPointSlide curSlide, Selection selection)
        {

            PowerPointBgEffectSlide effectSlide = EffectsLabUtil.GenerateEffectSlide(curSlide, selection, true);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.GothamBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public static void SepiaRemainderEffect(PowerPointSlide curSlide, Selection selection)
        {
            PowerPointBgEffectSlide effectSlide = EffectsLabUtil.GenerateEffectSlide(curSlide, selection, true);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.SepiaBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public static void GrayScaleBackgroundEffect(PowerPointSlide curSlide, Selection selection)
        {
            PowerPointBgEffectSlide effectSlide = EffectsLabUtil.GenerateEffectSlide(curSlide, selection, false);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.GrayScaleBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public static void BlackWhiteBackgroundEffect(PowerPointSlide curSlide, Selection selection)
        {
            PowerPointBgEffectSlide effectSlide = EffectsLabUtil.GenerateEffectSlide(curSlide, selection, false);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.BlackWhiteBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public static void GothamBackgroundEffect(PowerPointSlide curSlide, Selection selection)
        {
            PowerPointBgEffectSlide effectSlide = EffectsLabUtil.GenerateEffectSlide(curSlide, selection, false);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.GothamBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public static void SepiaBackgroundEffect(PowerPointSlide curSlide, Selection selection)
        {
            PowerPointBgEffectSlide effectSlide = EffectsLabUtil.GenerateEffectSlide(curSlide, selection, false);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.SepiaBackground();
            effectSlide.GetNativeSlide().Select();
        }
    }
}
