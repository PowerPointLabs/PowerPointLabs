using System;
using System.Collections.Generic;

using Microsoft.Office.Core;

using PowerPointLabs.Models;
using PowerPointLabs.SyncLab.ObjectFormats;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

using FillFormat = PowerPointLabs.SyncLab.ObjectFormats.FillFormat;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;
using Shapes = Microsoft.Office.Interop.PowerPoint.Shapes;

namespace PowerPointLabs.SyncLab
{
    /// <summary>
    /// Saves shapes into a PowerPointPresentation that exists in the background.
    /// The exact saved shapes may change in type but style will be retained.
    /// Eg: PlaceHolders are saved as Textboxes
    /// 
    /// 2013 only:
    /// We use a workabout to sync fill color, copying shapes gives the wrong fill
    /// We use a workabout to sync ArtisticEffecs
    /// </summary>
    public sealed class SyncLabShapeStorage : PowerPointPresentation
    {

        public const int FormatStorageSlide = 0;

        private int nextKey = 0;
        
        // only for 2013
        private readonly Dictionary<String, List<MsoPictureEffectType>> _backupArtisticEffects = 
            new Dictionary<string, List<MsoPictureEffectType>>();
        
        // only for 2013
        // need to sync all glow formats, syncing color alone resets transparency & radius
        // color must be synced first, it resets the transparency
        private readonly List<Format> _glowFormats = 
            new List<Format> 
            {
                new GlowColorFormat(),
                new GlowTransparencyFormat(),
                new GlowSizeFormat()
            };
        
        // only for 2013
        private readonly List<Format> _fillFormats =
            new List<Format>
            {
                new FillFormat()
            };

        private static readonly Lazy<SyncLabShapeStorage> StorageInstance =
            new Lazy<SyncLabShapeStorage>(() => new SyncLabShapeStorage());

        public static SyncLabShapeStorage Instance
        {
            get { return StorageInstance.Value; }
        }

        private SyncLabShapeStorage() : base()
        {
            Path = System.IO.Path.GetTempPath();
            Name = SyncLabText.StorageFileName;
            Open(withWindow: false, focus: false);
            ClearShapes();
        }

        public Shapes GetTemplateShapes()
        {
            return Slides[FormatStorageSlide].Shapes;
        }

        /// <summary>
        /// Saves shape in storage
        /// Returns a key to find the shape by,
        /// or null if the shape cannot be copied
        /// </summary>
        /// <param name="shape"></param>
        /// <param name="formats">Required for msoPlaceholder</param>
        /// <returns>identifier of copied shape</returns>
        public string CopyShape(Shape shape, Format[] formats)
        {
            Shape copiedShape = null;
            if (shape.Type == MsoShapeType.msoPlaceholder)
            {
                copiedShape = ShapeUtil.CopyMsoPlaceHolder(formats, shape, GetTemplateShapes());
            }
            else
            {
                try
                {
                    shape.Copy();
                    copiedShape = Slides[0].Shapes.Paste()[1];
                }
                catch
                {
                    copiedShape = null;
                }
            }

            if (copiedShape == null)
            {
                return null;
            }

            string shapeKey = nextKey.ToString();
            nextKey++;
            copiedShape.Name = shapeKey;

            #pragma warning disable 618
            if (Globals.ThisAddIn.IsApplicationVersion2013())
            {
                // sync glow, 2013 gives the wrong Glow.Fill color after copying the shape
                SyncFormats(shape, copiedShape, _glowFormats);
                // sync shape fill, 2013 gives the wrong fill color after copying the shape
                SyncFormats(shape, copiedShape, _fillFormats);
                
                // backup artistic effects for 2013
                // ForceSave() will make artistic effect permanent on the shapes for 2013 and no longer retrievable
                List<MsoPictureEffectType> extractedEffects = ArtisticEffectFormat.GetArtisticEffects(copiedShape);
                _backupArtisticEffects.Add(shapeKey, extractedEffects);
            }
            #pragma warning restore 618
            
            ForceSave();
            return shapeKey;
        }

        public Shape GetShape(string shapeKey)
        {
            Shapes shapes = Slides[0].Shapes;
            for (int i = 1; i <= shapes.Count; i++)
            {
                if (shapes[i].Name.Equals(shapeKey))
                {
                    Shape shape = shapes[i];
                    if (_backupArtisticEffects.ContainsKey(shapeKey))
                    {
                        // apply artistic effect from backup only when shape is retrieved, to reduce loading time of CopyShape
                        List<MsoPictureEffectType> extractedEffects = _backupArtisticEffects[shapeKey];
                        ArtisticEffectFormat.ClearArtisticEffects(shape);
                        ArtisticEffectFormat.ApplyArtisticEffects(shape, extractedEffects);
                    }
                    return shape;
                }
            }
            return null;
        }

        public void RemoveShape(string shapeKey)
        {
            int index = 1;
            Shapes shapes = Slides[0].Shapes;
            while (index <= shapes.Count)
            {
                if (shapes[index].Name.Equals(shapeKey))
                {
                    shapes[index].Delete();
                    _backupArtisticEffects.Remove(shapeKey);
                }
                else
                {
                    index++;
                }
            }
        }

        public void ForceSave()
        {
            Save();
            Close();
            Open(withWindow: false, focus: false);
        }

        public void ClearShapes()
        {
            while (SlideCount > 0)
            {
                GetSlide(1).Delete();
            }
            AddSlide();
            Slides[FormatStorageSlide].DeleteAllShapes();
            _backupArtisticEffects.Clear();
        }
        
        // Convenience method for syncing formats
        private void SyncFormats(Shape source, Shape destination, List<Format> formats)
        {
            foreach (var format in formats)
            {
                if (format.CanCopy(source))
                {
                    format.SyncFormat(source, destination);
                }
            }
        }

    }
}
