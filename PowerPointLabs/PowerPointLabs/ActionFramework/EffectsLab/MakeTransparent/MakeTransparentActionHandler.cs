using System.IO;
using System.Windows;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportActionRibbonId(EffectsLabText.MakeTransparentTag)]
    class MakeTransparentActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            var selection = this.GetCurrentSelection();

            if (selection.Type != PowerPoint.PpSelectionType.ppSelectionShapes)
            {
                MessageBox.Show("Please select at least 1 shape");
                return;
            }

            TransparentEffect(selection.ShapeRange);
        }

        private void TransparentEffect(PowerPoint.ShapeRange shapeRange)
        {
            foreach (PowerPoint.Shape shape in shapeRange)
            {
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    var subShapeRange = shape.Ungroup();
                    TransparentEffect(subShapeRange);
                    subShapeRange.Group();
                }
                else if (shape.Type == Office.MsoShapeType.msoPlaceholder)
                {
                    PlaceholderTransparencyHandler(shape);
                }
                else if (shape.Type == Office.MsoShapeType.msoPicture)
                {
                    PictureTransparencyHandler(shape);
                }
                else if (shape.Type == Office.MsoShapeType.msoLine)
                {
                    LineTransparencyHandler(shape);
                }
                else if (IsTransparentableShape(shape))
                {
                    ShapeTransparencyHandler(shape);
                }
            }
        }

        private bool IsTransparentableShape(PowerPoint.Shape shape)
        {
            return shape.Type == Office.MsoShapeType.msoAutoShape ||
                   shape.Type == Office.MsoShapeType.msoFreeform;
        }

        private void PictureTransparencyHandler(PowerPoint.Shape picture)
        {
            var rotation = picture.Rotation;

            picture.Rotation = 0;

            var tempPicPath = Path.Combine(Path.GetTempPath(), "tempPic.png");

            GraphicsUtil.ExportShape(picture, tempPicPath);

            var shapeHolder =
                this.GetCurrentSlide().Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    picture.Left,
                    picture.Top,
                    picture.Width,
                    picture.Height);

            var oriZOrder = picture.ZOrderPosition;

            picture.Delete();

            // move shape holder to original z-order
            while (shapeHolder.ZOrderPosition > oriZOrder)
            {
                shapeHolder.ZOrder(Office.MsoZOrderCmd.msoSendBackward);
            }

            shapeHolder.Line.Visible = Office.MsoTriState.msoFalse;
            shapeHolder.Fill.UserPicture(tempPicPath);
            shapeHolder.Fill.Transparency = 0.5f;

            shapeHolder.Rotation = rotation;

            File.Delete(tempPicPath);
        }

        private void PlaceholderTransparencyHandler(PowerPoint.Shape picture)
        {
            PictureTransparencyHandler(picture);
        }

        private void LineTransparencyHandler(PowerPoint.Shape shape)
        {
            shape.Line.Transparency = 0.5f;
        }

        private void ShapeTransparencyHandler(PowerPoint.Shape shape)
        {
            shape.Fill.Transparency = 0.5f;
            shape.Line.Transparency = 0.5f;
        }
    }
}
