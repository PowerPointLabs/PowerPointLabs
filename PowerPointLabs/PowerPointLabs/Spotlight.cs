using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    class Spotlight
    {
#pragma warning disable 0618
        public static float defaultSoftEdges = 10;
        public static float defaultTransparency = 0.25f;
        public static System.Drawing.Color defaultColor = Color.Black;
        public static Dictionary<String, float> softEdgesMapping = new Dictionary<string, float>
        {
            {"None", 0},
            {"1 Point", 1},
            {"2.5 Points", 2.5f},
            {"5 Points", 5},
            {"10 Points", 10},
            {"25 Points", 25},
            {"50 Points", 50}
        };
        public static void AddSpotlightEffect()
        {
            try
            {
                var currentSlide = PowerPointCurrentPresentationInfo.CurrentSlide as PowerPointSlide;
                PowerPoint.ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                
                var addedSlide = currentSlide.CreateSpotlightSlide() as PowerPointSpotlightSlide;
                List<PowerPoint.Shape> spotlightShapes = new List<PowerPoint.Shape>();

                addedSlide.DeleteShapesWithPrefix("SpotlightShape");
                foreach (PowerPoint.Shape spotShape in selectedShapes)
                {
                    addedSlide.DeleteShapesWithPrefix(spotShape.Name);
                    PreFormatShapeOnCurrentSlide(spotShape);
                    PowerPoint.Shape spotlightShape = addedSlide.CreateSpotlightShape(spotShape);
                    CreateSpotlightDuplicate(spotlightShape);
                    spotlightShapes.Add(spotlightShape);
                    PostFormatShapeOnCurrentSlide(currentSlide, spotShape);
                }
                
                addedSlide.PrepareForSpotlight();
                addedSlide.AddSpotlightEffect(spotlightShapes);
                currentSlide.DeleteShapesWithPrefix("SpotlightShape");
                PowerPointPresentation.Current.AddAckSlide();
            }
            catch (Exception e)
            {
                Logger.LogException(e, "SpotlightBtnClick");
                throw;
            }
        }

        private static void PreFormatShapeOnCurrentSlide(PowerPoint.Shape spotShape)
        {
            //Change color of shape to white. This is used later for creating spotlight shape
            spotShape.ShapeStyle = Office.MsoShapeStyleIndex.msoShapeStylePreset1;
            spotShape.Fill.ForeColor.RGB = 0xffffff;
            spotShape.Line.Visible = Office.MsoTriState.msoFalse;
            
            //Change color of text on shapes to white
            if (spotShape.HasTextFrame == Office.MsoTriState.msoTrue && spotShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                spotShape.TextFrame.TextRange.Font.Color.RGB = 0xffffff;

            //Deal with text on grouped shapes
            if (spotShape.Type == Office.MsoShapeType.msoGroup)
            {
                PowerPoint.ShapeRange shRange = spotShape.GroupItems.Range(1);
                foreach (PowerPoint.Shape sh in shRange)
                {
                    if (sh.HasTextFrame == Office.MsoTriState.msoTrue && sh.TextFrame.HasText == Office.MsoTriState.msoTrue)
                        sh.TextFrame.TextRange.Font.Color.RGB = 0xffffff;
                }
            }
        }

        private static void PostFormatShapeOnCurrentSlide(PowerPointSlide currentSlide, PowerPoint.Shape spotShape)
        {
            //Format selected shape on current slide
            spotShape.Fill.ForeColor.RGB = 0xaaaaaa;
            spotShape.Fill.Transparency = 0.7f;
            spotShape.Line.Visible = Office.MsoTriState.msoTrue;
            spotShape.Line.ForeColor.RGB = 0x000000;

            Utils.Graphics.MakeShapeViewTimeInvisible(spotShape, currentSlide);
        }
        
        private static void CreateSpotlightDuplicate(PowerPoint.Shape spotlightShape)
        {
            //Create hidden duplicate shape. This is needed for recreating spotlights 
            PowerPoint.Shape duplicateShape = spotlightShape.Duplicate()[1];
            duplicateShape.Visible = Office.MsoTriState.msoFalse;
            duplicateShape.Left = spotlightShape.Left;
            duplicateShape.Top = spotlightShape.Top;
        }
    }
}
