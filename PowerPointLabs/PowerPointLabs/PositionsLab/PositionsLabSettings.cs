using PowerPointLabs.ColorThemes.Extensions;
using PowerPointLabs.PositionsLab.Views;

namespace PowerPointLabs.PositionsLab
{
    public static class PositionsLabSettings
    {
        public enum AlignReferenceObject
        {
            Slide,
            SelectedShape,
            PowerpointDefaults
        }
        public static AlignReferenceObject AlignReference = AlignReferenceObject.SelectedShape;

        public enum DistributeReferenceObject
        {
            Slide,
            FirstShape,
            FirstTwoShapes,
            ExtremeShapes
        }
        public static DistributeReferenceObject DistributeReference = DistributeReferenceObject.FirstTwoShapes;

        public enum DistributeRadialReferenceObject
        {
            AtSecondShape,
            SecondThirdShape
        }
        public static DistributeRadialReferenceObject DistributeRadialReference = DistributeRadialReferenceObject.SecondThirdShape;

        public enum DistributeSpaceReferenceObject
        {
            ObjectBoundary,
            ObjectCenter
        }
        public static DistributeSpaceReferenceObject DistributeSpaceReference = DistributeSpaceReferenceObject.ObjectBoundary;

        public enum RadialShapeOrientationObject
        {
            Fixed,
            Dynamic
        }
        public static RadialShapeOrientationObject DistributeShapeOrientation = RadialShapeOrientationObject.Fixed;

        public enum GridAlignment
        {
            None,
            AlignLeft,
            AlignCenter,
            AlignRight
        }
        public static GridAlignment DistributeGridAlignment = GridAlignment.AlignLeft;
        public static float GridMarginTop = 5.0f;
        public static float GridMarginBottom = 5.0f;
        public static float GridMarginLeft = 5.0f;
        public static float GridMarginRight = 5.0f;

        public enum SwapReference
        {
            TopLeft,
            TopCenter,
            TopRight,
            MiddleLeft,
            MiddleCenter,
            MiddleRight,
            BottomLeft,
            BottomCenter,
            BottomRight
        }
        public static SwapReference SwapReferencePoint = SwapReference.MiddleCenter;
        public static bool IsSwapByClickOrder = false;

        public static RadialShapeOrientationObject ReorientShapeOrientation = RadialShapeOrientationObject.Fixed;

        public static void ShowAlignSettingsDialog()
        {
            AlignSettingsDialog dialog = new AlignSettingsDialog(AlignReference);
            dialog.DialogConfirmedHandler += OnAlignSettingsDialogConfirmed;
            dialog.ShowThematicDialog();
        }

        public static void ShowDistributeSettingsDialog()
        {
            DistributeSettingsDialog dialog = new DistributeSettingsDialog(DistributeReference, DistributeRadialReference, 
                                                                        DistributeSpaceReference, DistributeShapeOrientation);
            dialog.DialogConfirmedHandler += OnAlignSettingsDialogConfirmed;
            dialog.ShowThematicDialog();
        }

        public static void ShowReorderSettingsDialog()
        {
            ReorderSettingsDialog dialog = new ReorderSettingsDialog(IsSwapByClickOrder, SwapReferencePoint);
            dialog.DialogConfirmedHandler += OnReorderSettingsDialogConfirmed;
            dialog.ShowThematicDialog();
        }

        public static void ShowReorientSettingsDialog()
        {
            ReorientSettingsDialog dialog = new ReorientSettingsDialog(ReorientShapeOrientation);
            dialog.DialogConfirmedHandler += OnReorientSettingsDialogConfirmed;
            dialog.ShowThematicDialog();
        }

        private static void OnAlignSettingsDialogConfirmed(AlignReferenceObject alignReference)
        {
            AlignReference = alignReference;
        }

        private static void OnAlignSettingsDialogConfirmed(DistributeReferenceObject distributeReference,
                                                        DistributeRadialReferenceObject radialReference,
                                                        DistributeSpaceReferenceObject spaceReference,
                                                        RadialShapeOrientationObject orientationReference)
        {
            DistributeReference = distributeReference;
            DistributeRadialReference = radialReference;
            DistributeSpaceReference = spaceReference;
            DistributeShapeOrientation = orientationReference;
        }

        private static void OnReorderSettingsDialogConfirmed(bool isSwapByClickOrder, SwapReference swapReferencePoint)
        {
            IsSwapByClickOrder = isSwapByClickOrder;
            SwapReferencePoint = swapReferencePoint;
        }

        private static void OnReorientSettingsDialogConfirmed(RadialShapeOrientationObject reorientShapeOrientation)
        {
            ReorientShapeOrientation = reorientShapeOrientation;
        }
    }
}
