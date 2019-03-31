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
        public delegate void DialogConfirmedDelegate(MsoAutoShapeType newShapeType, MsoAnimEffect newAnimationType);
        public DialogConfirmedDelegate DialogConfirmedHandler { get; set; }

        private MsoAutoShapeType lastShapeType;
        private MsoAnimEffect lastAnimationType;

        public TooltipsLabSettingsDialogBox(MsoAutoShapeType shapeType, MsoAnimEffect animType)
        {
            lastShapeType = shapeType;
            lastAnimationType = animType;
            InitializeComponent();
            Initialize();
        }

        private void Initialize()
        {
            Array shapeTypes = Enum.GetValues(typeof(MsoAutoShapeType));
            Bitmap[] shapeBitmaps = ShapeTypesToBitmaps(shapeTypes, "Callout");
            for (int i = 0; i < shapeTypes.Length; i++)
            {
                if (shapeBitmaps[i] == null)
                {
                    continue;
                }
                TooltipsLabSettingsShapeEntry newEntry = new TooltipsLabSettingsShapeEntry(
                    (MsoAutoShapeType)shapeTypes.GetValue(i), shapeBitmaps[i]);
                shapeList.Items.Add(newEntry);
                if (newEntry.Type == lastShapeType)
                {
                    shapeList.SelectedItem = newEntry;
                    shapeList.ScrollIntoView(newEntry);
                }
            }

            Array animationTypes = Enum.GetValues(typeof(MsoAnimEffect));
            Bitmap dummyImage = Properties.Resources.AddSpotlightContext;
            for (int i = 0; i < animationTypes.Length; i++)
            {
                TooltipsLabSettingsAnimationEntry newEntry = new TooltipsLabSettingsAnimationEntry(
                    (MsoAnimEffect)animationTypes.GetValue(i), dummyImage);
                animationList.Items.Add(newEntry);
                if (newEntry.Type == lastAnimationType)
                {
                    animationList.SelectedItem = newEntry;
                    animationList.ScrollIntoView(newEntry);
                }
            }
        }

        private Bitmap[] ShapeTypesToBitmaps(Array types, string shapeType)
        {
            Shapes shapes = SyncFormatUtil.GetTemplateShapes();
            Bitmap[] bitmaps = new Bitmap[types.Length];
            for (int i = 0; i < types.Length; i++)
            {
                if (!((MsoAutoShapeType)types.GetValue(i)).ToString().Contains(shapeType))
                {
                    continue;
                }
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
            TooltipsLabSettingsAnimationEntry animationItem =
                animationList.SelectedItem as TooltipsLabSettingsAnimationEntry;
            if (shapeItem != null && animationItem != null)
            {
                DialogConfirmedHandler(shapeItem.Type, animationItem.Type);
            }
            Close();
        }

        #endregion

    }
}
