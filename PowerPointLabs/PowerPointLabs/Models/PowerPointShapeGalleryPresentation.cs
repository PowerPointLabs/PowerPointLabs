using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointShapeGalleryPresentation : PowerPointPresentation
    {
        private const string ShapeGalleryFileExtension = ".pptlabsshapes";
        private const string DuplicateShapeSuffixFormat = "(recovered shape {0})";
        
        private PowerPointSlide _defaultCategory;

        private readonly Dictionary<string, int> _categoryNameIndexMapper = new Dictionary<string, int>();

        # region Properties
        public string ShapeFolderPath { get; set; }
        # endregion

        # region Constructor
        public PowerPointShapeGalleryPresentation(string path, string name, string shapeFolderPath) : base(path, name)
        {
            ShapeFolderPath = shapeFolderPath;
        }
        public PowerPointShapeGalleryPresentation(Presentation presentation, string shapeFolderPath) : base(presentation)
        {
            ShapeFolderPath = shapeFolderPath;
        }
        # endregion

        # region API
        public void AddCategory(string name, bool setAsDefault = true)
        {
            if (_categoryNameIndexMapper.ContainsKey(name))
            {
                if (setAsDefault)
                {
                    _defaultCategory = Slides[_categoryNameIndexMapper[name] - 1];
                }

                return;
            }

            var newSlide = AddSlide(name: name);

            // ppLayoutBlank causes an error, so we use ppLayoutText instead and manually remove the
            // place holders
            newSlide.DeleteShapeWithRule(new Regex(@"^Title \d+$"));
            newSlide.DeleteShapeWithRule(new Regex(@"^Content Placeholder \d+$"));

            _categoryNameIndexMapper[name] = Slides.Count;

            if (setAsDefault)
            {
                _defaultCategory = newSlide;
            }
        }

        public void AddShape(Selection selection, string name)
        {
            selection.ShapeRange.Copy();

            var pastedShapeRange = _defaultCategory.Shapes.Paste();
            var pastedShape = pastedShapeRange[1];

            if (pastedShapeRange.Count > 1)
            {
                pastedShape = pastedShapeRange.Group();
            }

            pastedShape.Name = name;
        }

        public void AddShape(Selection selection, string category, string name)
        {
            selection.Copy();

            var categorySlide = Slides[_categoryNameIndexMapper[category]];
            var pastedShapeRange = categorySlide.Shapes.Paste();
            var pastedShape = pastedShapeRange[1];

            if (pastedShapeRange.Count > 1)
            {
                pastedShape = pastedShapeRange.Group();
            }

            pastedShape.Name = name;
        }

        public override void Close()
        {
            base.Close();

            RetrieveShapeGalleryFile();
        }

        public void CopyShape(string name)
        {
            var shapes = _defaultCategory.GetShapeWithName(name);

            if (shapes.Count != 1) return;
            
            shapes[0].Copy();
        }

        public override void Open(bool readOnly = false, bool untitled = false,
                                  bool withWindow = true, bool focus = true)
        {
            RetrievePptxFile();

            base.Open(readOnly, untitled, withWindow, focus);

            ConsistencyCheck();
        }

        public void RemoveCategory(string name)
        {
            if (_defaultCategory.Name == name)
            {
                _defaultCategory = null;
            }

            _categoryNameIndexMapper.Remove(name);

            RemoveSlide(name);
        }

        public void RemoveCategory(int index)
        {
            if (_defaultCategory.Name == Slides[index].Name)
            {
                _defaultCategory = null;
            }

            _categoryNameIndexMapper.Remove(Slides[index].Name);
            
            RemoveSlide(index);
        }

        public void RemoveShape(string name)
        {
            _defaultCategory.DeleteShapeWithName(name);
        }

        public void RenameShape(string oldName, string newName)
        {
            var shapes = _defaultCategory.GetShapeWithName(oldName);

            foreach (var shape in shapes)
            {
                shape.Name = newName;
            }
        }

        public void SetDefaultCategory(string name)
        {
            foreach (var slide in Slides)
            {
                if (slide.Name == name)
                {
                    _defaultCategory = slide;
                    break;
                }
            }
        }
        # endregion

        # region Helper Function
        private void ConsistencyCheck()
        {
            if (SlideCount < 1) return;

            // here we need to check 3 cases:
            // 1. self consistency check (if there are any duplicate names);
            // 2. more png than shapes inside pptx (shapes for short);
            // 3. more shapes than png.

            var shapeDuplicate = ConsistencyCheckSelf();

            var pngShapes = Directory.EnumerateFiles(ShapeFolderPath, "*.png").ToList();
            var shapeLost = ConsistencyCheckShapeToPng(pngShapes);
            var pngLost = ConsistencyCheckPngToShape(pngShapes);

            if (shapeDuplicate || shapeLost || pngLost)
            {
                MessageBox.Show(TextCollection.ShapeCorruptedError);
                Save();
            }
        }

        private bool ConsistencyCheckPngToShape(IEnumerable<string> pngShapes)
        {
            // if inconsistency is found, we delete the extra pngs
            var shapeLost = false;

            foreach (var pngShape in pngShapes)
            {
                var shapeName = System.IO.Path.GetFileNameWithoutExtension(pngShape);
                var found = Slides.Any(category => category.HasShapeWithSameName(shapeName));

                if (!found)
                {
                    shapeLost = true;
                    File.Delete(pngShape);
                }
            }

            return shapeLost;
        }

        private bool ConsistencyCheckSelf()
        {
            var shapeDuplicate = false;

            // if inconsistency is found, we keep all the shapes but:
            // 1. append "(recovered shape X)" to the shape name, X is the relative index
            // 2. export the shape as .png
            foreach (var category in Slides)
            {
                _categoryNameIndexMapper[category.Name] = category.Index;

                var shapeHash = new Dictionary<string, int>();
                var shapes = category.Shapes.Cast<Shape>().ToList();
                var duplicateShapeNames = new List<string>();

                foreach (var shape in shapes)
                {
                    if (shapeHash.Count == 0 ||
                        !shapeHash.ContainsKey(shape.Name))
                    {
                        shapeHash[shape.Name] = 1;
                    }
                    else
                    {
                        var index = (shapeHash[shape.Name] += 1);

                        // add to collection only if this shape is the first duplicate shape
                        if (index == 2)
                        {
                            duplicateShapeNames.Add(shape.Name);
                        }

                        RenameAndExportDuplicateShape(shape, index);
                    }
                }

                shapeDuplicate = shapeDuplicate || duplicateShapeNames.Count > 0;

                foreach (var lastShapeName in duplicateShapeNames)
                {
                    var lastShapePath = ShapeFolderPath + @"\" + lastShapeName + ".png";
                    var lastShape = category.GetShapeWithName(lastShapeName)[0];

                    File.Delete(lastShapePath);
                    RenameAndExportDuplicateShape(lastShape, 1);
                }
            }

            return shapeDuplicate;
        }

        private bool ConsistencyCheckShapeToPng(List<string> pngShapes)
        {
            // if inconsistency is found, we delete the extra shape
            var shapeLost = false;

            foreach (var category in Slides)
            {
                // this is to handle the case when user deletes the .png image manually but
                // ShapeGallery.pptx isn't updated
                var shapeCnt = 1;

                while (shapeCnt <= category.Shapes.Count)
                {
                    var shape = category.Shapes[shapeCnt];
                    var shapePath = ShapeFolderPath + @"\" + shape.Name + ".png";

                    if (!pngShapes.Contains(shapePath))
                    {
                        shape.Delete();
                        shapeLost = true;
                    }
                    else
                    {
                        shapeCnt++;
                    }
                }
            }

            return shapeLost;
        }

        private void RetrievePptxFile()
        {
            var shapeGalleryFileName = FullName.Replace(".pptx", ShapeGalleryFileExtension);

            if (File.Exists(shapeGalleryFileName))
            {
                File.SetAttributes(shapeGalleryFileName, FileAttributes.Normal);
                File.Move(shapeGalleryFileName, FullName);
            }

            // to reduce the chance that user opens the shape gallery file, we make the pptx file hidden
            File.SetAttributes(FullName, FileAttributes.Hidden);
        }

        private void RetrieveShapeGalleryFile()
        {
            // set the file as a visible readonly .pptlabsshapes file.
            var shapeGalleryFileName = FullName.Replace(".pptx", ShapeGalleryFileExtension);

            File.Move(FullName, shapeGalleryFileName);

            File.SetAttributes(shapeGalleryFileName, FileAttributes.Normal);
            File.SetAttributes(shapeGalleryFileName, FileAttributes.ReadOnly);
        }

        private void RenameAndExportDuplicateShape(Shape shape, int index)
        {
            shape.Name += string.Format(DuplicateShapeSuffixFormat, index);

            var shapeExportPath = ShapeFolderPath + @"\" + shape.Name + ".png";

            shape.Export(shapeExportPath, PpShapeFormat.ppShapeFormatPNG, ExportMode: PpExportMode.ppScaleXY);
        }
        # endregion
    }
}
