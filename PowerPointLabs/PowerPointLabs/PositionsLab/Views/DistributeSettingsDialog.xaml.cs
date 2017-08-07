using System.Windows;

namespace PowerPointLabs.PositionsLab.Views
{
    /// <summary>
    /// Interaction logic for DistributeSettingsDialog.xaml
    /// </summary>
    public partial class DistributeSettingsDialog
    {
        public delegate void DialogConfirmedDelegate(PositionsLabSettings.DistributeReferenceObject distributeReference,
                                                    PositionsLabSettings.DistributeRadialReferenceObject radialReference,
                                                    PositionsLabSettings.DistributeSpaceReferenceObject spaceReference,
                                                    PositionsLabSettings.RadialShapeOrientationObject orientationReference);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        public DistributeSettingsDialog()
        {
            InitializeComponent();
        }

        public DistributeSettingsDialog(PositionsLabSettings.DistributeReferenceObject distributeReference,
                                        PositionsLabSettings.DistributeRadialReferenceObject radialReference,
                                        PositionsLabSettings.DistributeSpaceReferenceObject spaceReference,
                                        PositionsLabSettings.RadialShapeOrientationObject orientationReference)
            : this()
        {
            switch (distributeReference)
            {
                case PositionsLabSettings.DistributeReferenceObject.Slide:
                    distributeToSlideButton.IsChecked = true;
                    break;
                case PositionsLabSettings.DistributeReferenceObject.FirstShape:
                    distributeToFirstShapeButton.IsChecked = true;
                    break;
                case PositionsLabSettings.DistributeReferenceObject.FirstTwoShapes:
                    distributeToFirstTwoShapesButton.IsChecked = true;
                    break;
                case PositionsLabSettings.DistributeReferenceObject.ExtremeShapes:
                    distributeToExtremeShapesButton.IsChecked = true;
                    break;
            }

            switch (radialReference)
            {
                case PositionsLabSettings.DistributeRadialReferenceObject.AtSecondShape:
                    distributeAtSecondShapeButton.IsChecked = true;
                    break;
                case PositionsLabSettings.DistributeRadialReferenceObject.SecondThirdShape:
                    distributeToSecondThirdShapeButton.IsChecked = true;
                    break;
            }

            switch (spaceReference)
            {
                case PositionsLabSettings.DistributeSpaceReferenceObject.ObjectBoundary:
                    distributeByBoundariesButton.IsChecked = true;
                    break;
                case PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter:
                    distributeByShapeCenterButton.IsChecked = true;
                    break;
            }

            switch (orientationReference)
            {
                case PositionsLabSettings.RadialShapeOrientationObject.Fixed:
                    distributeShapeOrientationFixedButton.IsChecked = true;
                    break;
                case PositionsLabSettings.RadialShapeOrientationObject.Dynamic:
                    distributeShapeOrientationDynamicButton.IsChecked = true;
                    break;
            }
        }
        
        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabSettings.DistributeReferenceObject distributeReference;
            PositionsLabSettings.DistributeRadialReferenceObject radialReference;
            PositionsLabSettings.DistributeSpaceReferenceObject spaceReference;
            PositionsLabSettings.RadialShapeOrientationObject orientationReference;
            
            // Checks for boundary reference
            if (distributeToSlideButton.IsChecked.GetValueOrDefault())
            {
                distributeReference = PositionsLabSettings.DistributeReferenceObject.Slide;
            }
            else if (distributeToFirstShapeButton.IsChecked.GetValueOrDefault())
            {
                distributeReference = PositionsLabSettings.DistributeReferenceObject.FirstShape;
            }
            else if (distributeToFirstTwoShapesButton.IsChecked.GetValueOrDefault())
            {
                distributeReference = PositionsLabSettings.DistributeReferenceObject.FirstTwoShapes;
            }
            else
            {
                distributeReference = PositionsLabSettings.DistributeReferenceObject.ExtremeShapes;
            }

            // Checks for radial boundary reference
            if (distributeAtSecondShapeButton.IsChecked.GetValueOrDefault())
            {
                radialReference = PositionsLabSettings.DistributeRadialReferenceObject.AtSecondShape;
            }
            else
            {
                radialReference = PositionsLabSettings.DistributeRadialReferenceObject.SecondThirdShape;
            }

            // Checks for space calculation reference
            if (distributeByBoundariesButton.IsChecked.GetValueOrDefault())
            {
                spaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectBoundary;
            }
            else
            {
                spaceReference = PositionsLabSettings.DistributeSpaceReferenceObject.ObjectCenter;
            }

            // Checks for radial shape orientation
            if (distributeShapeOrientationFixedButton.IsChecked.GetValueOrDefault())
            {
                orientationReference = PositionsLabSettings.RadialShapeOrientationObject.Fixed;
            }
            else
            {
                orientationReference = PositionsLabSettings.RadialShapeOrientationObject.Dynamic;
            }

            DialogConfirmedHandler(distributeReference, radialReference, spaceReference, orientationReference);
            Close();
        }
    }
}
