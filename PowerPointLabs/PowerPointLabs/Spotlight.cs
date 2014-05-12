using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using PowerPointLabs.Models;
using Office = Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs
{
    class Spotlight
    {
        public static float defaultSoftEdges = 10;
        public static float defaultTransparency = 0.7f;
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
                var currentSlide = PowerPointPresentation.CurrentSlide as PowerPointSlide;
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
                AddAckSlide();
            }
            catch (Exception e)
            {
                //LogException(e, "SpotlightBtnClick");
                throw;
            }
        }

        public static void ReloadSpotlightEffect()
        {
            try
            {
                var currentSlide = PowerPointPresentation.CurrentSlide as PowerPointSpotlightSlide;
                if (currentSlide.isSpotlightSlide())
                {
                    PowerPoint.Shape spotlightPicture = null;
                    PowerPoint.Shape indicatorShape = null;
                    List<PowerPoint.Shape> spotlightShapes = new List<PowerPoint.Shape>();

                    foreach (PowerPoint.Shape sh in currentSlide.Shapes)
                    {
                        if (sh.Name.Equals("SpotlightShape1"))
                        {
                            spotlightPicture = sh;
                        }
                        else if (sh.Name.Contains("SpotlightShape"))
                        {
                            spotlightShapes.Add(sh);
                        }
                        else if (sh.Name.Contains("PPTLabsIndicator"))
                        {
                            indicatorShape = sh;
                        }
                    }

                    if (spotlightPicture == null || spotlightShapes.Count == 0)
                    {
                        System.Windows.Forms.MessageBox.Show("The spotlight effect cannot be recreated for the current slide.\nPlease click on the Create Spotlight button to create a new spotlight.", "Error");
                    }
                    else
                    {
                        spotlightPicture.Delete();
                        if (indicatorShape != null)
                            indicatorShape.Delete();

                        foreach (PowerPoint.Shape sh in spotlightShapes)
                        {
                            sh.Visible = Office.MsoTriState.msoTrue;
                            CreateSpotlightDuplicate(sh);
                        }

                        currentSlide.PrepareForSpotlight();
                        currentSlide.AddSpotlightEffect(spotlightShapes);
                        AddAckSlide();
                    }
                }
                else
                {
                    System.Windows.Forms.MessageBox.Show("The current slide is not a spotlight slide added by the PowerPointLabs plugin", "Error");
                }
            }
            catch (Exception e)
            {
                //LogException(e, "ReloadSpotlightButtonClick");
                throw;
            }
        }

        private static void PreFormatShapeOnCurrentSlide(PowerPoint.Shape spotShape)
        {
            spotShape.ShapeStyle = Office.MsoShapeStyleIndex.msoShapeStylePreset1;
            spotShape.Fill.ForeColor.RGB = 0xffffff;
            spotShape.Line.Visible = Office.MsoTriState.msoFalse;
            
            //Change color of text on shapes to white
            if (spotShape.HasTextFrame == Office.MsoTriState.msoTrue && spotShape.TextFrame.HasText == Office.MsoTriState.msoTrue)
                spotShape.TextFrame.TextRange.Font.Color.RGB = 0xffffff;

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
            spotShape.Fill.ForeColor.RGB = 0xaaaaaa;
            spotShape.Fill.Transparency = 0.7f;
            spotShape.Line.Visible = Office.MsoTriState.msoTrue;
            spotShape.Line.ForeColor.RGB = 0x000000;

            PowerPoint.Effect effectAppear = null;
            PowerPoint.Effect effectDisappear = null;

            PowerPoint.Sequence sequence = currentSlide.TimeLine.MainSequence;
            effectAppear = sequence.AddEffect(spotShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectAppear.Timing.Duration = 0;
            effectAppear.MoveTo(1);

            effectDisappear = sequence.AddEffect(spotShape, PowerPoint.MsoAnimEffect.msoAnimEffectAppear, PowerPoint.MsoAnimateByLevel.msoAnimateLevelNone, PowerPoint.MsoAnimTriggerType.msoAnimTriggerWithPrevious);
            effectDisappear.Exit = Office.MsoTriState.msoTrue;
            effectDisappear.Timing.Duration = 0;
            effectDisappear.MoveTo(2);
        }
        
        private static void CreateSpotlightDuplicate(PowerPoint.Shape spotlightShape)
        {
            //Create hidden duplicate shape. This is needed for recreating spotlights 
            PowerPoint.Shape duplicateShape = spotlightShape.Duplicate()[1];
            duplicateShape.Visible = Office.MsoTriState.msoFalse;
            duplicateShape.Left = spotlightShape.Left;
            duplicateShape.Top = spotlightShape.Top;
        }

        private static void AddAckSlide()
        {
            try
            {
                PowerPointSlide lastSlide = PowerPointPresentation.Slides.Last();
                if (!lastSlide.isAckSlide())
                {
                    lastSlide.CreateAckSlide();
                }
            }
            catch (Exception e)
            {
                //LogException(e, "AddAckSlide");
                throw;
            }
        }
    }
}
