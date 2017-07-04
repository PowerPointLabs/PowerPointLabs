using System.Drawing;

using PowerPointLabs.Views;

namespace PowerPointLabs.EffectsLab
{
    internal static class EffectsLabSpotlightSettings
    {
        public static void OpenSpotlightSettingsDialog()
        {
            SpotlightSettingsDialogBox dialog = new SpotlightSettingsDialogBox(Spotlight.defaultTransparency,
                                                                                    Spotlight.defaultSoftEdges,
                                                                                    Spotlight.defaultColor);
            dialog.SettingsHandler += SpotlightPropertiesEdited;
            dialog.ShowDialog();
        }

        public static void SpotlightPropertiesEdited(float newTransparency, float newSoftEdge, Color newColor)
        {
            Spotlight.defaultTransparency = newTransparency;
            Spotlight.defaultSoftEdges = newSoftEdge;
            Spotlight.defaultColor = newColor;
        }
    }
}
