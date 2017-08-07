using System;
using System.Windows;

using PowerPointLabs.ActionFramework.Common.Attribute;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Interface;
using PowerPointLabs.CropLab;
using PowerPointLabs.TextCollection;

using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.ActionFramework.EffectsLab
{
    [ExportActionRibbonId(EffectsLabText.MagnifyTag)]
    class MagnifyActionHandler : ActionHandler
    {
        protected override void ExecuteAction(string ribbonId)
        {
            this.StartNewUndoEntry();

            var selection = this.GetCurrentSelection();

            PowerPoint.ShapeRange shapeRange;

            try
            {
                shapeRange = selection.ShapeRange;
            }
            catch (Exception)
            {
                MessageBox.Show("Please select an area to magnify.");

                return;
            }

            if (shapeRange.Count > 1 || shapeRange[1].Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
            {
                MessageBox.Show("Only one magnify area is allowed.");

                return;
            }

            try
            {
                var croppedShape = CropToShape.Crop(this.GetCurrentSlide(), selection, isInPlace: true, handleError: false);

                MagnifyGlassEffect(croppedShape, 1.4f);
            }
            catch (Exception e)
            {
                var errorMessage = e.Message;
                errorMessage = errorMessage.Replace("Crop To Shape", "Magnify");

                MessageBox.Show(errorMessage);
            }
        }

        private void MagnifyGlassEffect(PowerPoint.Shape shape, float ratio)
        {
            var delta = 0.5f * (ratio - 1);

            shape.Left -= delta * shape.Width;
            shape.Top -= delta * shape.Height;

            shape.Width *= ratio;
            shape.Height *= ratio;

            shape.Shadow.Visible = Office.MsoTriState.msoTrue;
            shape.Shadow.Style = Office.MsoShadowStyle.msoShadowStyleOuterShadow;
            shape.Shadow.Transparency = 0.6f;
            shape.Shadow.Size = 102f;
            shape.Shadow.Blur = 5;
            shape.Shadow.OffsetX = 0;
            shape.Shadow.OffsetY = 2f;

            shape.ThreeD.BevelTopType = Office.MsoBevelType.msoBevelCircle;
            shape.ThreeD.BevelTopInset = 15;
            shape.ThreeD.BevelTopDepth = 3;
            shape.ThreeD.BevelBottomType = Office.MsoBevelType.msoBevelNone;
            shape.ThreeD.PresetLighting = Office.MsoLightRigType.msoLightRigBalanced;
            shape.ThreeD.LightAngle = 145;

            shape.LockAspectRatio = Office.MsoTriState.msoTrue;
        }
    }
}
