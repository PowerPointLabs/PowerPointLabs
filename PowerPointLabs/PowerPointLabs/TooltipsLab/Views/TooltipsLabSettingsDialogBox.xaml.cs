using System;
using System.Drawing;
using System.Windows;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.SyncLab;
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

        private MsoAutoShapeType lastSelectedShapeType;
        private MsoAnimEffect lastSelectedAnimType;

        public TooltipsLabSettingsDialogBox(MsoAutoShapeType selectedShapeType, MsoAnimEffect selectedAnimType)
        {
            lastSelectedShapeType = selectedShapeType;
            lastSelectedAnimType = selectedAnimType;
            InitializeComponent();
            Initialize();
        }

        private void Initialize()
        {
            Array shapeTypes = Enum.GetValues(typeof(MsoAutoShapeType));
            Bitmap[] shapeBitmaps = ShapeTypesToBitmaps(shapeTypes, TooltipsLabConstants.CalloutNameSubstring);
            for (int i = 0; i < shapeTypes.Length; i++)
            {
                if (shapeBitmaps[i] == null)
                {
                    continue;
                }
                TooltipsLabSettingsShapeEntry newEntry = new TooltipsLabSettingsShapeEntry(
                    (MsoAutoShapeType)shapeTypes.GetValue(i), shapeBitmaps[i]);
                shapeList.Items.Add(newEntry);
                if (newEntry.Type == lastSelectedShapeType)
                {
                    shapeList.SelectedItem = newEntry;
                    shapeList.ScrollIntoView(newEntry);
                }
            }

            for (int i = 0; i < TooltipsLabConstants.AnimationEffects.Length; i++)
            {
                TooltipsLabSettingsAnimationEntry newEntry = new TooltipsLabSettingsAnimationEntry(
                    TooltipsLabConstants.AnimationEffects[i],
                    TooltipsLabConstants.AnimationImages[i]);
                animationList.Items.Add(newEntry);
                if (newEntry.Type == lastSelectedAnimType)
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
                        TooltipsLabConstants.DisplayImageSize.Width,
                        TooltipsLabConstants.DisplayImageSize.Height);
                    ShapeUtil.FormatCalloutToDefaultStyle(shape);
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
