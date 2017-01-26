using System;
using System.Diagnostics;

using Microsoft.Office.Core;

namespace PowerPointLabs.SyncLab.ObjectFormats
{
    class FillFormat : ObjectFormat
    {
#pragma warning disable 0618
        #region Properties
        
        private readonly Microsoft.Office.Interop.PowerPoint.ColorFormat backColor;
        private readonly Microsoft.Office.Interop.PowerPoint.ColorFormat foreColor;
        private readonly float gradientAngle;
        private readonly MsoGradientColorType gradientColorType;
        private readonly float gradientDegree;
        private readonly GradientStops gradientStops;
        private readonly MsoGradientStyle gradientStyle;
        private readonly int gradientVariant;
        private readonly MsoPatternType pattern;
        private readonly PictureEffects pictureEffects;
        private readonly MsoPresetGradientType presetGradientType;
        private readonly MsoPresetTexture presetTexture;
        private readonly MsoTriState rotateWithObject;
        private readonly MsoTextureAlignment textureAlignment;
        private readonly float textureHorizontalScale;
        private readonly string textureName;
        private readonly float textureOffsetX;
        private readonly float textureOffsetY;
        private readonly MsoTriState textureTile;
        private readonly MsoTextureType textureType;
        private readonly float textureVerticalScale;
        private readonly float transparency;
        private readonly MsoFillType type;
        private readonly MsoTriState visible;
        #endregion

        public FillFormat(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            
            this.displayText = "Fill";
            this.displayImage = Utils.Graphics.ShapeToImage(shape);
            Microsoft.Office.Interop.PowerPoint.FillFormat format = shape.Fill;
            /*
            this.backColor = format.BackColor;
            this.foreColor = format.ForeColor;
            this.transparency = format.Transparency;
            this.visible = format.Visible;
            this.type = format.Type;
            this.rotateWithObject = format.RotateWithObject;

            switch (format.Type)
            {
                case MsoFillType.msoFillBackground:
                    // Nothing to copy
                    break;
                case MsoFillType.msoFillGradient:
                    // ??
                    break;
                case MsoFillType.msoFillMixed:
                    // ??
                    break;
                case MsoFillType.msoFillPatterned:
                    this.pattern = format.Pattern;
                    break;
                case MsoFillType.msoFillPicture:
                    // ??
                    break;
                case MsoFillType.msoFillSolid:
                    // Nothing to copy
                    break;
                case MsoFillType.msoFillTextured:
                    this.textureAlignment = format.TextureAlignment;
                    this.textureHorizontalScale = format.TextureHorizontalScale;
                    this.textureOffsetX = format.TextureOffsetX;
                    this.textureOffsetY = format.TextureOffsetY;
                    this.textureTile = format.TextureTile;
                    this.textureType = format.TextureType;
                    this.textureVerticalScale = format.TextureVerticalScale;
                    break;

            }

            this.pattern = format.Pattern;
            this.textureAlignment = format.TextureAlignment;
            this.textureHorizontalScale = format.TextureHorizontalScale;
            this.textureOffsetX = format.TextureOffsetX;
            this.textureOffsetY = format.TextureOffsetY;
            this.textureTile = format.TextureTile;
            this.textureType = format.TextureType;
            this.textureVerticalScale = format.TextureVerticalScale;

            /*
        Debug.WriteLine("========== CREATING FILL OBJECT ==========");
            try
            {
                this.backColor = format.BackColor;
                Debug.WriteLine("BackColor: " + backColor.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!BackColor: " + ex.Message);
            }
            try
            {
                this.foreColor = format.ForeColor;
                Debug.WriteLine("ForeColor: " + foreColor.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!Foreolor: " + ex.Message);
            }
            try
            {
                this.gradientAngle = format.GradientAngle;
                Debug.WriteLine("GradientAngle: " + gradientAngle.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!GradientAngle: " + ex.Message);
            }
            try
            {
                this.gradientColorType = format.GradientColorType;
                Debug.WriteLine("GradientColorType: " + gradientColorType.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!GradientColorType: " + ex.Message);
            }
            try
            {
                this.gradientDegree = format.GradientDegree;
                Debug.WriteLine("GradientDegree: " + gradientDegree.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!GradientDegree: " + ex.Message);
            }
            try
            {
                this.gradientStops = format.GradientStops;
                Debug.WriteLine("GradientStops: " + gradientStops.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!GradientStops: " + ex.Message);
            }
            try
            {
                this.gradientStyle = format.GradientStyle;
                Debug.WriteLine("GradientStyle: " + gradientStyle.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!GradientStyle: " + ex.Message);
            }
            try
            {
                this.gradientVariant = format.GradientVariant;
                Debug.WriteLine("GradientVariant: " + gradientVariant.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!GradientVariant: " + ex.Message);
            }
            try
            {
                this.pattern = format.Pattern;
                Debug.WriteLine("Pattern: " + pattern.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!Pattern: " + ex.Message);
            }
            try
            {
                this.pictureEffects = format.PictureEffects;
                Debug.WriteLine("PictureEffects: " + pictureEffects.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!PictureEffects: " + ex.Message);
            }
            try
            {
                this.presetGradientType = format.PresetGradientType;
                Debug.WriteLine("PresetGradientType: " + presetGradientType.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!PresetGradientType: " + ex.Message);
            }
            try
            {
                this.presetTexture = format.PresetTexture;
                Debug.WriteLine("PresetTexture: " + presetTexture.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!PresetTexture: " + ex.Message);
            }
            try
            {
                this.rotateWithObject = format.RotateWithObject;
                Debug.WriteLine("RotateWithObject: " + rotateWithObject.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!RotateWithObject: " + ex.Message);
            }
            try
            {
                this.textureAlignment = format.TextureAlignment;
                Debug.WriteLine("TextureAlignment: " + textureAlignment.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!TextureAlignment: " + ex.Message);
            }
            try
            {
                this.textureHorizontalScale = format.TextureHorizontalScale;
                Debug.WriteLine("TextureHorizontalScale: " + textureHorizontalScale.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!TextureHorizontalScale: " + ex.Message);
            }
            try
            {
                this.textureName = format.TextureName;
                Debug.WriteLine("TextureName: " + textureName.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!TextureName: " + ex.Message);
            }
            try
            {
                this.textureOffsetX = format.TextureOffsetX;
                Debug.WriteLine("TextureOffsetX: " + textureOffsetX.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!TextureOffsetX: " + ex.Message);
            }
            try
            {
                this.textureOffsetY = format.TextureOffsetY;
                Debug.WriteLine("TextureOffsetY: " + textureOffsetY.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!TextureOffsetY: " + ex.Message);
            }
            try
            {
                this.textureTile = format.TextureTile;
                Debug.WriteLine("TextureTile: " + textureTile.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!TextureTile: " + ex.Message);
            }
            try
            {
                this.textureType = format.TextureType;
                Debug.WriteLine("TextureType: " + textureType.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!TextureType: " + ex.Message);
            }
            try
            {
                this.textureVerticalScale = format.TextureVerticalScale;
                Debug.WriteLine("TextureVerticalScale: " + textureVerticalScale.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!TextureVerticalScale: " + ex.Message);
            }
            try
            {
                this.transparency = format.Transparency;
                Debug.WriteLine("Transparency: " + transparency.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!Transparency: " + ex.Message);
            }
            try
            {
                this.type = format.Type;
                Debug.WriteLine("Type: " + type.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!Type: " + ex.Message);
            }
            try
            {
                this.visible = format.Visible;
                Debug.WriteLine("Visible: " + visible.ToString());
            }
            catch (Exception ex)
            {
                Debug.WriteLine("!Visible: " + ex.Message);
            }
            */

  
            this.backColor = format.BackColor;
            this.foreColor = format.ForeColor;
            this.gradientAngle = format.GradientAngle;
            this.gradientColorType = format.GradientColorType;
            this.gradientDegree = format.GradientDegree;
            this.gradientStops = format.GradientStops;
            this.gradientStyle = format.GradientStyle;
            this.gradientVariant = format.GradientVariant;
            this.pattern = format.Pattern;
            this.pictureEffects = format.PictureEffects;
            this.presetGradientType = format.PresetGradientType;
            this.presetTexture = format.PresetTexture;
            this.rotateWithObject = format.RotateWithObject;
            this.textureAlignment = format.TextureAlignment;
            this.textureHorizontalScale = format.TextureHorizontalScale;
            this.textureName = format.TextureName;
            this.textureOffsetX = format.TextureOffsetX;
            this.textureOffsetY = format.TextureOffsetY;
            this.textureTile = format.TextureTile;
            this.textureType = format.TextureType;
            this.textureVerticalScale = format.TextureVerticalScale;
            this.transparency = format.Transparency;
            this.type = format.Type;
            this.visible = format.Visible;
        }

        public override void ApplyTo(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            /*
            Microsoft.Office.Interop.PowerPoint.FillFormat format = shape.Fill;

            format.BackColor = this.backColor;
            format.ForeColor = this.foreColor;
            format.Transparency = this.transparency;
            format.Visible = this.visible;
            //format.Type = this.type;
            format.RotateWithObject = this.rotateWithObject;

            switch (format.Type)
            {
                case MsoFillType.msoFillBackground:
                    // Nothing to copy
                    break;
                case MsoFillType.msoFillGradient:
                    // ??
                    break;
                case MsoFillType.msoFillMixed:
                    // ??
                    break;
                case MsoFillType.msoFillPatterned:
                    format.Pattern = this.pattern;
                    break;
                case MsoFillType.msoFillPicture:
                    // ??
                    break;
                case MsoFillType.msoFillSolid:
                    // Nothing to copy
                    break;
                case MsoFillType.msoFillTextured:
                    format.textureAlignment = this.TextureAlignment;
                    format.textureHorizontalScale = this.TextureHorizontalScale;
                    format.textureOffsetX = this.TextureOffsetX;
                    format.textureOffsetY = this.TextureOffsetY;
                    format.textureTile = this.TextureTile;
                    format.textureType = this.TextureType;
                    format.textureVerticalScale = this.TextureVerticalScale;
                    break;
            }*/
            
            Microsoft.Office.Interop.PowerPoint.FillFormat format = shape.Fill;
            format.BackColor = this.backColor;
            format.ForeColor = this.foreColor;
            format.Transparency = this.transparency;
            format.Visible = this.visible;
            format.RotateWithObject = this.rotateWithObject;
            format.TextureAlignment = this.textureAlignment;
            format.TextureHorizontalScale = this.textureHorizontalScale;
            format.TextureOffsetX = this.textureOffsetX;
            format.TextureOffsetY = this.textureOffsetY;
            format.TextureVerticalScale = this.textureVerticalScale;
           
        }
    }
}