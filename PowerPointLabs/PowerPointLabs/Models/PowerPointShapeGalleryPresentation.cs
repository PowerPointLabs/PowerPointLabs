using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.Utils;

namespace PowerPointLabs.Models
{
    class PowerPointShapeGalleryPresentation : PowerPointPresentation
    {
        private const string ShapeGalleryFileExtension = ".pptlabsshapes";
        private const string DuplicateShapeSuffixFormat = "(recovered shape {0})";

        private const int MaxUndoAmount = 20;
        
        private PowerPointSlide _defaultCategory;

        private readonly Dictionary<string, int> _categoryNameIndexMapper = new Dictionary<string, int>();

        # region Properties
        public List<string> Categories { get; private set; }
        public string DefaultCategory
        {
            get { return _defaultCategory.Name; }
            set
            {
                if (_categoryNameIndexMapper.ContainsKey(value))
                {
                    _defaultCategory = Slides[_categoryNameIndexMapper[value] - 1];
                }
            }
        }
        public bool IsImportedFile { get; set; }
        # endregion

        # region Constructor
        public PowerPointShapeGalleryPresentation(string path, string name) : base(path, name) {}
        public PowerPointShapeGalleryPresentation(Presentation presentation) : base(presentation) {}
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
            Categories.Add(name);

            if (setAsDefault)
            {
                _defaultCategory = newSlide;
            }

            Save();
            ActionProtection();
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

            Save();
            ActionProtection();
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

            Save();
            ActionProtection();
        }

        public override void Close()
        {
            base.Close();

            _categoryNameIndexMapper.Clear();
            RetrieveShapeGalleryFile();
        }

        public void CopyShape(string name)
        {
            // copy a shape with name in the default category
            var shapes = _defaultCategory.GetShapeWithName(name);

            if (shapes.Count != 1) return;
            
            shapes[0].Copy();
        }

        public bool HasCategory(string name)
        {
            return _categoryNameIndexMapper.ContainsKey(name);
        }

        public void MoveShape(string name, string categoryName)
        {
            if (!_categoryNameIndexMapper.ContainsKey(categoryName)) return;

            // move a shape with name from default category to another category
            var shapes = _defaultCategory.GetShapeWithName(name);
            var destCategory = Slides[_categoryNameIndexMapper[categoryName]];

            if (shapes.Count != 1) return;

            shapes[0].Cut();
            destCategory.Shapes.Paste();

            Save();
            ActionProtection();
        }

        public override bool Open(bool readOnly = false, bool untitled = false,
                                  bool withWindow = true, bool focus = true)
        {
            RetrievePptxFile();

            // if we can't even open the file, return false
            if (!base.Open(readOnly, untitled, withWindow, focus))
            {
                return false;
            }

            PrepareCategories();
            
            return ConsistencyCheck();
        }

        public void RemoveCategory(string name)
        {
            if (_defaultCategory.Name == name)
            {
                _defaultCategory = null;
            }

            _categoryNameIndexMapper.Remove(name);
            Categories.Remove(name);

            RemoveSlide(name);

            Save();
            ActionProtection();
        }

        public void RemoveCategory(int index)
        {
            if (_defaultCategory.Name == Slides[index].Name)
            {
                _defaultCategory = null;
            }

            _categoryNameIndexMapper.Remove(Slides[index].Name);
            Categories.RemoveAt(index);
            
            RemoveSlide(index);

            Save();
            ActionProtection();
        }

        public void RemoveCategory()
        {
            var index = _categoryNameIndexMapper[_defaultCategory.Name];
            
            Categories.RemoveAt(index);

            RemoveSlide(index);

            _categoryNameIndexMapper.Remove(_defaultCategory.Name);

            Save();
            ActionProtection();
        }

        public void RemoveShape(string name)
        {
            _defaultCategory.DeleteShapeWithName(name);
            
            Save();
            ActionProtection();
        }

        public void RenameShape(string oldName, string newName)
        {
            var shapes = _defaultCategory.GetShapeWithName(oldName);

            foreach (var shape in shapes)
            {
                shape.Name = newName;
            }

            Save();
            ActionProtection();
        }

        public void RenameCategory(string newName)
        {
            var index = _categoryNameIndexMapper[_defaultCategory.Name];

            _categoryNameIndexMapper.Remove(_defaultCategory.Name);

            _defaultCategory.Name = newName;
            _categoryNameIndexMapper[newName] = index;

            Save();
            ActionProtection();
        }
        # endregion

        # region Helper Function
        private void ActionProtection()
        {
            for (var i = 0; i < MaxUndoAmount; i ++)
            {
                Presentation.Slides[1].Background.Fill.BackColor = Presentation.Slides[1].Background.Fill.BackColor;
            }
        }

        private bool ConsistencyCheck()
        {
            // if there's no slide, the file is always valid
            if (SlideCount < 1) return true;

            // here we need to check 3 cases:
            // 1. self consistency check (if there are any duplicate names);
            // 2. more png than shapes inside pptx (shapes for short);
            // 3. more shapes than png.

            var shapeDuplicate = ConsistencyCheckSelf();
            var shapeLost = false;
            var pngLost = false;

            foreach (var category in Slides)
            {
                var shapeFolderPath = Path + @"\" + category.Name;

                // check if we have a corresponding category directory in the Path
                ConsistencyCheckCategorySlideToLocal(category);

                var pngShapes = Directory.EnumerateFiles(shapeFolderPath, "*.png").ToList();

                // critical: OR with itself at the end to avoid early truncate
                shapeLost = ConsistencyCheckShapeToPng(pngShapes, category) || shapeLost;
                pngLost = ConsistencyCheckPngToShape(pngShapes, category) || pngLost;
            }

            var categoryInShapeGalleryLost = ConsistencyCheckCategoryLocalToSlide();

            Save();

            if (shapeDuplicate || shapeLost || categoryInShapeGalleryLost)
            {
                MessageBox.Show(TextCollection.ShapeCorruptedError);

                return false;
            }

            if (pngLost && !IsImportedFile)
            {
                MessageBox.Show(TextCollection.ShapeCorruptedError);

                return false;
            }

            return true;
        }

        private bool ConsistencyCheckCategoryLocalToSlide()
        {
            var categoriesOnDisk = Directory.EnumerateDirectories(Path).ToList();
            var categoryLost = false;

            foreach (var categoryPath in categoriesOnDisk)
            {
                var categoryName = new DirectoryInfo(categoryPath).Name;

                if (Slides.All(category => category.Name != categoryName))
                {
                    FileDir.DeleteFolder(categoryPath);

                    categoryLost = true;
                }
            }

            return categoryLost;
        }

        private void ConsistencyCheckCategorySlideToLocal(PowerPointSlide category)
        {
            var categoryFolderPath = Path + @"\" + category.Name;

            // the category is some how lost on the disk, regenerate the category
            if (!Directory.Exists(categoryFolderPath))
            {
                // create the directory
                Directory.CreateDirectory(categoryFolderPath);
                // since shape reconstruction will be taken care of during ConsistencyCheckShapeToPng(),
                // we do not need to generate the shapes here
            }
        }

        private bool ConsistencyCheckPngToShape(IEnumerable<string> pngShapes, PowerPointSlide category)
        {
            // if inconsistency is found, we delete the extra pngs
            var shapeLost = false;

            foreach (var pngShape in pngShapes)
            {
                var shapeName = System.IO.Path.GetFileNameWithoutExtension(pngShape);
                var found = category.HasShapeWithSameName(shapeName);

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
                var shapeHash = new Dictionary<string, int>();
                var shapes = category.Shapes.Cast<Shape>().ToList();
                var duplicateShapeNames = new List<string>();

                var shapeFolderPath = Path + @"\" + category.Name;

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

                        RenameAndExportDuplicateShape(shape, index, shapeFolderPath);
                    }
                }

                shapeDuplicate = shapeDuplicate || duplicateShapeNames.Count > 0;

                foreach (var lastShapeName in duplicateShapeNames)
                {
                    var lastShapePath = shapeFolderPath + @"\" + lastShapeName + ".png";
                    var lastShape = category.GetShapeWithName(lastShapeName)[0];

                    File.Delete(lastShapePath);
                    RenameAndExportDuplicateShape(lastShape, 1, shapeFolderPath);
                }
            }

            return shapeDuplicate;
        }

        private bool ConsistencyCheckShapeToPng(List<string> pngShapes, PowerPointSlide category)
        {
            // if inconsistency is found, we export the extra shape to .png
            var shapeLost = false;
            var shapeFolderPath = Path + @"\" + category.Name;

            // this is to handle 2 cases:
            // 1. user deleted the .png shape accidentally;
            // 2. the file is imported
            foreach (Shape shape in category.Shapes)
            {
                var shapePath = shapeFolderPath + @"\" + shape.Name + ".png";

                if (!pngShapes.Contains(shapePath))
                {
                    Graphics.ExportShape(shape, shapePath);
                    shapeLost = true;
                }
            }

            return shapeLost;
        }

        private void PrepareCategories()
        {
            if (SlideCount < 1) return;

            if (Categories != null)
            {
                Categories.Clear();
            }
            else
            {
                Categories = new List<string>();
            }

            // record each slide in index-name mapper
            foreach (var category in Slides)
            {
                _categoryNameIndexMapper[category.Name] = category.Index;
                Categories.Add(category.Name);

                if (category.Index == 0)
                {
                    _defaultCategory = category;
                }
            }
        }

        private void RetrievePptxFile()
        {
            var shapeGalleryFileName = FullName.Replace(".pptx", ShapeGalleryFileExtension);

            if (File.Exists(shapeGalleryFileName))
            {
                File.SetAttributes(shapeGalleryFileName, FileAttributes.Normal);
                File.Move(shapeGalleryFileName, FullName);

                // to reduce the chance that user opens the shape gallery file, we make the pptx file hidden
                File.SetAttributes(FullName, FileAttributes.Hidden);
            }
        }

        private void RetrieveShapeGalleryFile()
        {
            // set the file as a visible readonly .pptlabsshapes file.
            var shapeGalleryFileName = FullName.Replace(".pptx", ShapeGalleryFileExtension);

            Trace.TraceInformation("FullName = " + FullName + ", Name = " + shapeGalleryFileName);

            File.Move(FullName, shapeGalleryFileName);

            File.SetAttributes(shapeGalleryFileName, FileAttributes.Normal);
            File.SetAttributes(shapeGalleryFileName, FileAttributes.ReadOnly);
        }

        private void RenameAndExportDuplicateShape(Shape shape, int index, string shapeFolderPath)
        {
            shape.Name += string.Format(DuplicateShapeSuffixFormat, index);

            var shapeExportPath = shapeFolderPath + @"\" + shape.Name + ".png";

            Graphics.ExportShape(shape, shapeExportPath);
        }
        # endregion
    }
}
