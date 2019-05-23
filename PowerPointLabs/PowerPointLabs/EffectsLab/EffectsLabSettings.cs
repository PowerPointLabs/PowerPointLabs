using System.Collections.Generic;
using System.Drawing;

using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.EffectsLab.Views;
using PowerPointLabs.TextCollection;

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
                case EffectsLabText.BlurrinessFeatureSelected:
                    isTint = IsTintSelected;
                    customPercentage = CustomPercentageSelected;
                    break;
                case EffectsLabText.BlurrinessFeatureRemainder:
                    isTint = IsTintRemainder;
                    customPercentage = CustomPercentageRemainder;
                    break;
                case EffectsLabText.BlurrinessFeatureBackground:
                    isTint = IsTintBackground;
                    customPercentage = CustomPercentageBackground;
                    break;
                default:
                    return;
            }

            BlurSettingsDialogBox dialog = new BlurSettingsDialogBox(feature, isTint, customPercentage);
            dialog.DialogConfirmedHandler += OnBlurSettingsDialogConfirmed;
            dialog.ShowThematicDialog();
        }

        public static void ShowSpotlightSettingsDialog()
        {
            SpotlightSettingsDialogBox dialog = new SpotlightSettingsDialogBox(SpotlightTransparency, SpotlightSoftEdges, SpotlightColor);
            dialog.DialogConfirmedHandler += OnSpotlightSettingsDialogConfirmed;
            dialog.ShowThematicDialog();
        }

        private static void OnBlurSettingsDialogConfirmed(string feature, bool isTint, int customPercentage)
        {
            switch (feature)
            {
                case EffectsLabText.BlurrinessFeatureSelected:
                    IsTintSelected = isTint;
                    CustomPercentageSelected = customPercentage;
                    break;
                case EffectsLabText.BlurrinessFeatureRemainder:
                    IsTintRemainder = isTint;
                    CustomPercentageRemainder = customPercentage;
                    break;
                case EffectsLabText.BlurrinessFeatureBackground:
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
