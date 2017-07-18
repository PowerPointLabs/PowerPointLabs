using System.Drawing;

using PowerPointLabs.EffectsLab.Views;

namespace PowerPointLabs.EffectsLab
{
    internal static class EffectsLabSpotlightSettings
    {
        public static void OpenSpotlightSettingsDialog()
        {
            SpotlightSettingsDialogBox dialog = new SpotlightSettingsDialogBox(Spotlight.transparency,
                                                                                    Spotlight.softEdges,
                                                                                    Spotlight.color);
            dialog.DialogConfirmedHandler += SpotlightPropertiesEdited;
            dialog.ShowDialog();
        }

        public static void SpotlightPropertiesEdited(float newTransparency, float newSoftEdge, Color newColor)
        {
            Spotlight.transparency = newTransparency;
            Spotlight.softEdges = newSoftEdge;
            Spotlight.color = newColor;
        }
    }
}
