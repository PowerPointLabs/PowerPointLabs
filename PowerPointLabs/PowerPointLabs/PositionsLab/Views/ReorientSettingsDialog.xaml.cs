using System.Windows;

namespace PowerPointLabs.PositionsLab.Views
{
    /// <summary>
    /// Interaction logic for ReorientSettingsDialog.xaml
    /// </summary>
    public partial class ReorientSettingsDialog
    {
        public delegate void DialogConfirmedDelegate(PositionsLabSettings.RadialShapeOrientationObject reorientShapeOrientation);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        public ReorientSettingsDialog()
        {
            InitializeComponent();
        }

        public ReorientSettingsDialog(PositionsLabSettings.RadialShapeOrientationObject reorientShapeOrientation)
            : this()
        {
            switch (reorientShapeOrientation)
            {
                case PositionsLabSettings.RadialShapeOrientationObject.Fixed:
                    reorientShapeOrientationFixedButton.IsChecked = true;
                    break;
                case PositionsLabSettings.RadialShapeOrientationObject.Dynamic:
                    reorientShapeOrientationDynamicButton.IsChecked = true;
                    break;
            }
        }

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            PositionsLabSettings.RadialShapeOrientationObject reorientShapeOrientation;

            if (reorientShapeOrientationFixedButton.IsChecked.GetValueOrDefault())
            {
                reorientShapeOrientation = PositionsLabSettings.RadialShapeOrientationObject.Fixed;
            }
            else
            {
                reorientShapeOrientation = PositionsLabSettings.RadialShapeOrientationObject.Dynamic;
            }

            DialogConfirmedHandler(reorientShapeOrientation);
            Close();
        }
    }
}
