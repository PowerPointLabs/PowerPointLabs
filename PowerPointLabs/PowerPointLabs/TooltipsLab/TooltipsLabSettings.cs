using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.TooltipsLab.Views;

using MsoAutoShapeType = Microsoft.Office.Core.MsoAutoShapeType;

namespace PowerPointLabs.TooltipsLab
{
    internal static class TooltipsLabSettings
    {
        public static MsoAutoShapeType ShapeType = MsoAutoShapeType.msoShapeRoundedRectangularCallout;
        public static MsoAnimEffect AnimationType = MsoAnimEffect.msoAnimEffectFade;
        public static bool IsUseFrameAnimation = false;

        public static void ShowSettingsDialog()
        {
            TooltipsLabSettingsDialogBox dialog = new TooltipsLabSettingsDialogBox(ShapeType, AnimationType);
            dialog.DialogConfirmedHandler += OnSettingsDialogConfirmed;
            dialog.ShowDialog();
        }

        private static void OnSettingsDialogConfirmed(MsoAutoShapeType newShapeType, MsoAnimEffect newAnimationType)
        {
            ShapeType = newShapeType;
        }
    }
}