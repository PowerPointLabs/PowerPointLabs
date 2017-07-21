using System.Collections.Generic;
using System.Drawing;

using PowerPointLabs.EffectsLab.Views;

namespace PowerPointLabs.EffectsLab
{
    internal static class EffectsLabSettings
    {
        public static Dictionary<string, float> SpotlightSoftEdgesMapping = new Dictionary<string, float>
        {
            {"None", 0},
            {"1 Point", 1},
            {"2.5 Points", 2.5f},
            {"5 Points", 5},
            {"10 Points", 10},
            {"25 Points", 25},
            {"50 Points", 50}
        };
        public static float SpotlightSoftEdges = 10.0f;
        public static float SpotlightTransparency = 0.25f;
        public static Color SpotlightColor = Color.Black;

        public static bool IsTintSelected = false;
        public static bool IsTintRemainder = false;
        public static bool IsTintBackground = false;

        public static int CustomPercentageSelected = 30;
        public static int CustomPercentageRemainder = 30;
        public static int CustomPercentageBackground = 30;

        public static void ShowBlurSettingsDialog(string feature)
        {
            bool isTint;
            int customPercentage;

            switch (feature)
            {
                case TextCollection1.EffectsLabBlurrinessFeatureSelected:
                    isTint = IsTintSelected;
                    customPercentage = CustomPercentageSelected;
                    break;
                case TextCollection1.EffectsLabBlurrinessFeatureRemainder:
                    isTint = IsTintRemainder;
                    customPercentage = CustomPercentageRemainder;
                    break;
                case TextCollection1.EffectsLabBlurrinessFeatureBackground:
                    isTint = IsTintBackground;
                    customPercentage = CustomPercentageBackground;
                    break;
                default:
                    return;
            }

            BlurSettingsDialogBox dialog = new BlurSettingsDialogBox(feature, isTint, customPercentage);
            dialog.DialogConfirmedHandler += OnBlurSettingsDialogConfirmed;
            dialog.ShowDialog();
        }

        public static void ShowSpotlightSettingsDialog()
        {
            SpotlightSettingsDialogBox dialog = new SpotlightSettingsDialogBox(SpotlightTransparency, SpotlightSoftEdges, SpotlightColor);
            dialog.DialogConfirmedHandler += OnSpotlightSettingsDialogConfirmed;
            dialog.ShowDialog();
        }

        private static void OnBlurSettingsDialogConfirmed(string feature, bool isTint, int customPercentage)
        {
            switch (feature)
            {
                case TextCollection1.EffectsLabBlurrinessFeatureSelected:
                    IsTintSelected = isTint;
                    CustomPercentageSelected = customPercentage;
                    break;
                case TextCollection1.EffectsLabBlurrinessFeatureRemainder:
                    IsTintRemainder = isTint;
                    CustomPercentageRemainder = customPercentage;
                    break;
                case TextCollection1.EffectsLabBlurrinessFeatureBackground:
                    IsTintBackground = isTint;
                    CustomPercentageBackground = customPercentage;
                    break;
                default:
                    break;
            }
        }

        private static void OnSpotlightSettingsDialogConfirmed(float spotlightTransparency, float spotlightSoftEdges, Color spotlightColor)
        {
            SpotlightTransparency = spotlightTransparency;
            SpotlightSoftEdges = spotlightSoftEdges;
            SpotlightColor = spotlightColor;
        }
    }
}
