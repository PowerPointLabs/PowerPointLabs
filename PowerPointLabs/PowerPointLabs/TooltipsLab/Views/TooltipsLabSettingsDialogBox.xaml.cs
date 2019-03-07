using System;
using System.Drawing;
using System.Windows;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.SyncLab;
using PowerPointLabs.SyncLab.ObjectFormats;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace PowerPointLabs.TooltipsLab.Views
{
    /// <summary>
    /// Interaction logic for TooltipsLabSettingsDialogBox.xaml
    /// </summary>
    public partial class TooltipsLabSettingsDialogBox
    {
        public delegate void DialogConfirmedDelegate(MsoAutoShapeType newShapeType);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        private MsoAutoShapeType lastShapeType;

        public TooltipsLabSettingsDialogBox()
        {
            InitializeComponent();
            Initialize();
        }
        
        public TooltipsLabSettingsDialogBox(MsoAutoShapeType defaultShapeType)
            : this()
        {
            lastShapeType = defaultShapeType;
        }

        private void Initialize()
        {
            Array shapeTypes = Enum.GetValues(typeof(MsoAutoShapeType));
            Bitmap[] shapeBitmaps = ShapeTypesToBitmaps(shapeTypes);
            for (int i = 0; i < shapeTypes.Length; i++)
            {
                if (shapeBitmaps[i] == null)
                {
                    continue;
                }
                shapeList.Items.Add(new TooltipsLabSettingsShapeEntry((MsoAutoShapeType)shapeTypes.GetValue(i), shapeBitmaps[i]));
            }

            Array animationTypes = Enum.GetValues(typeof(MsoAnimEffect));
            //Bitmap[] animationBitmaps;
        }

        private Bitmap[] ShapeTypesToBitmaps(Array types)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Bitmap[] bitmaps = new Bitmap[types.Length];
            for (int i = 0; i < types.Length; i++)
            {
                try
                {
                    Shape shape = shapes.AddShape(
                        (MsoAutoShapeType)types.GetValue(i), 0, 0,
                        SyncFormatConstants.DisplayImageSize.Width,
                        SyncFormatConstants.DisplayImageSize.Height);
                    bitmaps[i] = new Bitmap(GraphicsUtil.ShapeToBitmap(shape));
                    shape.Delete();
                }
                catch
                {

                }
            }
            return bitmaps;
        }

        #region EventHandlers

        private void OkButton_Click(object sender, RoutedEventArgs e)
        {
            TooltipsLabSettingsShapeEntry shapeItem = shapeList.SelectedItem as TooltipsLabSettingsShapeEntry;
            DialogConfirmedHandler(shapeItem.Type);
            Close();
        }

        #endregion

    }
}
