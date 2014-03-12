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
        public static void AddSpotlightEffect()
        {
            try
            {
                var currentSlide = PowerPointPresentation.CurrentSlide;
                PowerPoint.ShapeRange selectedShapes = Globals.ThisAddIn.Application.ActiveWindow.Selection.ShapeRange;
                
                var addedSlide = currentSlide.CreateSpotlightSlide() as PowerPointSpotlightSlide;
                List<PowerPoint.Shape> spotlightShapes = new List<PowerPoint.Shape>();

                addedSlide.DeleteShapesWithPrefix("SpotlightShape");
                foreach (PowerPoint.Shape spotShape in selectedShapes)
                {
                    addedSlide.DeleteShapesWithPrefix(spotShape.Name);
                    PowerPoint.Shape spotlightShape = addedSlide.CreateSpotlightShape(spotShape);
                    CreateSpotlightDuplicate(spotlightShape);
                    spotlightShapes.Add(spotlightShape);
                    spotShape.Delete();
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
                    }

                    if (spotlightPicture == null || spotlightShapes.Count == 0)
                    {
                        System.Windows.Forms.MessageBox.Show("The spotlight effect cannot be recreated for the current slide.\nPlease click on the Create Spotlight button to create a new spotlight.", "Error");
                    }
                    else
                    {
                        spotlightPicture.Delete();

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
