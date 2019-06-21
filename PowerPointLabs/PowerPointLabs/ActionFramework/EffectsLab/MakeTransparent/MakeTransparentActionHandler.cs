using System.IO;
using System.Windows;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;
using Office = Microsoft.Office.Core;


namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportActionRibbonId(EffectsLabText.MakeTransparentTag)]
    class MakeTransparentActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            Selection selection = this.GetCurrentSelection();

            if (selection.Type != PpSelectionType.ppSelectionShapes)
            {
                MessageBoxUtil.Show(TextCollection.EffectsLabText.ErrorSelectAtLeastOneShape);
                return;
            }

            TransparentEffect(selection.ShapeRange);
        }

        private void TransparentEffect(ShapeRange shapeRange)
        {
            foreach (Shape shape in shapeRange)
            {
                if (shape.Type == Office.MsoShapeType.msoGroup)
                {
                    ShapeRange subShapeRange = shape.Ungroup();
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

        private bool IsTransparentableShape(Shape shape)
        {
            return shape.Type == Office.MsoShapeType.msoAutoShape ||
                   shape.Type == Office.MsoShapeType.msoFreeform;
        }

        private void PictureTransparencyHandler(Shape picture)
        {
            float rotation = picture.Rotation;

            picture.Rotation = 0;

            string tempPicPath = Path.Combine(Path.GetTempPath(), "tempPic.png");

            GraphicsUtil.ExportShape(picture, tempPicPath);

            Shape shapeHolder =
                this.GetCurrentSlide().Shapes.AddShape(
                    Office.MsoAutoShapeType.msoShapeRectangle,
                    picture.Left,
                    picture.Top,
                    picture.Width,
                    picture.Height);

            int oriZOrder = picture.ZOrderPosition;

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

        private void PlaceholderTransparencyHandler(Shape picture)
        {
            PictureTransparencyHandler(picture);
        }

        private void LineTransparencyHandler(Shape shape)
        {
            shape.Line.Transparency = 0.5f;
        }

        private void ShapeTransparencyHandler(Shape shape)
        {
            shape.Fill.Transparency = 0.5f;
            shape.Line.Transparency = 0.5f;
        }
    }
}
