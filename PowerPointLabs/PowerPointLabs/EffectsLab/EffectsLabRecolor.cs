using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.Models;
using PowerPointLabs.Views;

namespace PowerPointLabs.EffectsLab
{
    internal static class EffectsLabRecolor
    {
        public static void GreyScaleRemainderEffect(PowerPointSlide curSlide, Selection selection)
        {
            var effectSlide = GenerateEffectSlide(curSlide, selection, true);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.GreyScaleBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public static void BlackWhiteRemainderEffect(PowerPointSlide curSlide, Selection selection)
        {
            var effectSlide = GenerateEffectSlide(curSlide, selection, true);
            
            if (effectSlide == null)
            {
                return;
            }

            effectSlide.BlackWhiteBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public static void GothamRemainderEffect(PowerPointSlide curSlide, Selection selection)
        {

            var effectSlide = GenerateEffectSlide(curSlide, selection, true);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.GothamBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public static void SepiaRemainderEffect(PowerPointSlide curSlide, Selection selection)
        {
            var effectSlide = GenerateEffectSlide(curSlide, selection, true);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.SepiaBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public static void GreyScaleBackgroundEffect(PowerPointSlide curSlide, Selection selection)
        {
            var effectSlide = GenerateEffectSlide(curSlide, selection, false);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.GreyScaleBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public static void BlackWhiteBackgroundEffect(PowerPointSlide curSlide, Selection selection)
        {
            var effectSlide = GenerateEffectSlide(curSlide, selection, false);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.BlackWhiteBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public static void GothamBackgroundEffect(PowerPointSlide curSlide, Selection selection)
        {
            var effectSlide = GenerateEffectSlide(curSlide, selection, false);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.GothamBackground();
            effectSlide.GetNativeSlide().Select();
        }

        public static void SepiaBackgroundEffect(PowerPointSlide curSlide, Selection selection)
        {
            var effectSlide = GenerateEffectSlide(curSlide, selection, false);

            if (effectSlide == null)
            {
                return;
            }

            effectSlide.SepiaBackground();
            effectSlide.GetNativeSlide().Select();
        }

        internal static PowerPointBgEffectSlide GenerateEffectSlide(PowerPointSlide curSlide, Selection selection, bool generateOnRemainder)
        {
            PowerPointSlide dupSlide = null;

            try
            {
                var shapeRange = selection.ShapeRange;

                if (shapeRange.Count != 0)
                {
                    dupSlide = curSlide.Duplicate();
                }

                shapeRange.Cut();

                var effectSlide = PowerPointBgEffectSlide.BgEffectFactory(curSlide.GetNativeSlide(), generateOnRemainder);

                if (dupSlide != null)
                {
                    if (generateOnRemainder)
                    {
                        dupSlide.Delete();
                    }
                    else
                    {
                        dupSlide.MoveTo(curSlide.Index);
                        curSlide.Delete();
                    }
                }

                return effectSlide;
            }
            catch (InvalidOperationException e)
            {
                MessageBox.Show(e.Message);
                return null;
            }
            catch (COMException)
            {
                if (dupSlide != null)
                {
                    dupSlide.Delete();
                }

                MessageBox.Show("Please select at least 1 shape");
                return null;
            }
            catch (Exception e)
            {
                if (dupSlide != null)
                {
                    dupSlide.Delete();
                }

                ErrorDialogWrapper.ShowDialog("Error", e.Message, e);
                return null;
            }
        }
    }
}
