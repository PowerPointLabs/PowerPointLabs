using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace PowerPointLabs.ELearningLab.Utility
{
    public class ShapeUtility
    {
#pragma warning disable 0618
        /// <summary>
        /// Insert default callout box shape to slide. 
        /// Precondition: shape with shapeName must not exist in slide before applying the method
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="shapeName">shapeName is a string with format "PPTL_{tagNo}_Callout" to be associated 
        /// with generated callout shape.</param>
        /// <param name="calloutText">Content in Callout Shape</param>
        /// <returns>generated callout shape</returns>
        public static Shape InsertDefaultCalloutBoxToSlide(PowerPointSlide slide, string shapeName, string calloutText)
        {
            float slideWidth = PowerPointPresentation.Current.SlideWidth;
            float slideHeight = PowerPointPresentation.Current.SlideHeight;

            Shape calloutBox = slide.Shapes.AddShape(MsoAutoShapeType.msoShapeRoundedRectangularCallout, 10, 10, 100, 10);
            calloutBox.Name = shapeName;
            calloutBox.TextFrame.TextRange.Text = calloutText;
            calloutBox.Left = 10;
            calloutBox.Top = 10;
            calloutBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            calloutBox.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentLeft;
            ShapeUtil.FormatCalloutToDefaultStyle(calloutBox);

            return calloutBox;
        }

        /// <summary>
        /// Insert default caption box to slide
        /// Precondition: shape with shapeName must not exist in slide before applying the method
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="shapeName"></param>
        /// <param name="captionText"></param>
        /// <returns>the generated default caption box</returns>
        public static Shape InsertDefaultCaptionBoxToSlide(PowerPointSlide slide, string shapeName, string captionText)
        {
            float slideWidth = PowerPointPresentation.Current.SlideWidth;
            float slideHeight = PowerPointPresentation.Current.SlideHeight;

            Shape captionBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, slideHeight - 100,
                slideWidth, 100);
            captionBox.Name = shapeName;
            captionBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            captionBox.TextFrame.TextRange.Text = captionText;
            captionBox.TextFrame.WordWrap = MsoTriState.msoTrue;
            captionBox.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
            captionBox.TextFrame.TextRange.Font.Size = 12;
            captionBox.Fill.BackColor.RGB = 0;
            captionBox.Fill.Transparency = 0.2f;
            captionBox.TextFrame.TextRange.Font.Color.RGB = 0xffffff;

            captionBox.Top = slideHeight - captionBox.Height;
            return captionBox;
        }

        /// <summary>
        /// Insert shape which is copied from `templatedShape` to slide
        /// Precondition: shapeName must not exist in slide before applying the method
        /// </summary>
        /// <param name="slide"></param>
        /// <param name="templatedShape">Shape whose format is to be copied over</param>
        /// <param name="shapeName"></param>
        /// <param name="text"></param>
        /// <returns>the copied shape</returns>
        public static Shape InsertTemplatedShapeToSlide(PowerPointSlide slide, Shape templatedShape, string shapeName, string text)
        {
            float slideWidth = PowerPointPresentation.Current.SlideWidth;
            float slideHeight = PowerPointPresentation.Current.SlideHeight;

            // templatedShape and its associated animations are duplicated
            templatedShape.Copy();
            Shape copiedShape = slide.Shapes.Paste()[1];
            copiedShape.Name = shapeName;
            // copy shape to the default callout / caption position
            if (StringUtility.ExtractFunctionFromString(copiedShape.Name) == ELearningLabText.CalloutIdentifier)
            {
                copiedShape.Left = 10;
                copiedShape.Top = 10;
                copiedShape.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentLeft;
            }
            else if (StringUtility.ExtractFunctionFromString(copiedShape.Name) == ELearningLabText.CaptionIdentifier)
            {
                copiedShape.Left = 0;
                copiedShape.Top = slideHeight - 100;
                copiedShape.Width = slideWidth;
                copiedShape.Height = 100;
                copiedShape.Top = slideHeight - copiedShape.Height;
                copiedShape.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
            }

            copiedShape.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            copiedShape.TextFrame.TextRange.Text = text;
            copiedShape.TextFrame.WordWrap = MsoTriState.msoTrue;         
            // remove associated animation with copiedShape because we only want the shape to be copied.
            slide.RemoveAnimationsForShape(copiedShape);
            if (StringUtility.ExtractFunctionFromString(copiedShape.Name) == ELearningLabText.CaptionIdentifier)
            {
                copiedShape.Top = slideHeight - copiedShape.Height;
            }
            return copiedShape;
        }

        /// <summary>
        /// Replace original text in `shape` with `text`
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="text"></param>
        /// <returns></returns>
        public static Shape ReplaceTextForShape(Shape shape, string text)
        {
            shape.TextFrame.TextRange.Text = text;
            return shape;
        }

        public static Shape InsertSelfExplanationTextBoxToSlide(PowerPointSlide slide, string shapeName, string text)
        {
            float slideWidth = PowerPointPresentation.Current.SlideWidth;
            float slideHeight = PowerPointPresentation.Current.SlideHeight;

            Shape captionBox = slide.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 0, 0,
                slideWidth, 100);
            captionBox.Name = shapeName;
            captionBox.TextFrame.AutoSize = PpAutoSize.ppAutoSizeShapeToFitText;
            captionBox.TextFrame.TextRange.Text = text;
            captionBox.TextFrame.WordWrap = MsoTriState.msoTrue;
            captionBox.TextEffect.Alignment = MsoTextEffectAlignment.msoTextEffectAlignmentCentered;
            captionBox.TextFrame.TextRange.Font.Size = 12;
            captionBox.Fill.BackColor.RGB = 0xffffff;
            captionBox.Fill.Transparency = 0.2f;
            captionBox.TextFrame.TextRange.Font.Color.RGB = 0;
            captionBox.Visible = MsoTriState.msoFalse;

            return captionBox;
        }
    }
}
