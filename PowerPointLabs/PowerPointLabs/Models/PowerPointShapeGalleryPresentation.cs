using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

using Microsoft.Office.Interop.PowerPoint;
using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;

using Office = Microsoft.Office.Core;

namespace PowerPointLabs.Models
{
    class PowerPointShapeGalleryPresentation : PowerPointPresentation
    {
        /************************************************************************
         * Some General Concerns
         * 
         * 1. Be careful when using PowerPointPresentation.Slides property. The
         * implementation requires O(n) time to access an item, instead of O(1).
         * Therefore, when features in PowerPointSlide is not required, access
         * slides via PowerPointPresentation.Presentation.Slides;
         ************************************************************************/
        private const string CategoryNameBoxSearchPattern = "[Cc]ategory: *([^<>:\"/\\\\|?*]+)";
        private const string CategoryNameFormat = "Category: {0}";
        private const string DefaultSlideNameSearchPattern = @"[Ss]lide ?\d+";
        private const string DuplicateShapeSuffixFormat = "(duplicate shape {0})";
        private const string GroupSelectionNameFormat = "Group {0} Seq_{1}";
        private const string GroupSelectionNamePattern = @"^Group ([\w\s]+) Seq_(\d+)$";
        private const string NameSearchPattern = @"^Group {0} Seq_(\d+)$|^{1}$";
        private const string NameExtractionPatternFormat = @"^Group ({0}(?: \d+)*) Seq_\d+$|^({1}(?: \d+)*)$";
        private const string ShapeGalleryFileExtension = ".pptlabsshapes";
        private const string UntitledCategoryNameFormat = "Untitled Category {0}";

        private const int MaxUndoAmount = 20;
        
        private PowerPointSlide _defaultCategory;
        private readonly List<Shape> _categoryNameBoxCollection = new List<Shape>();

        # region Properties
        public List<string> Categories { get; private set; }

        public string DefaultCategory
        {
            get
            {
                if (_defaultCategory == null)
                {
                    return null;
                }

                return _defaultCategory.Name;
            }
            set
            {
                FindCategoryIndex(value, true);
            }
        }

        public string ImportToCategory { get; set; }

        public bool IsImportedFile { get; set; }
        # endregion

        # region Constructor
        public PowerPointShapeGalleryPresentation(string path, string name) : base(path, name)
        {
            Categories = new List<string>();
        }
        public PowerPointShapeGalleryPresentation(Presentation presentation) : base(presentation)
        {
            Categories = new List<string>();
        }
        # endregion

        # region API
        /// <summary>
        /// Remember to lock/release clipboard when using from clipboard!
        /// </summary>
        public void AddCategory(string name, bool setAsDefault = true, bool fromClipBoard = false)
        {
            int index = FindCategoryIndex(name, setAsDefault);

            // the category already exists
            if (index != -1)
            {
                return;
            }

            // ppLayoutBlank causes an error, so we use ppLayoutText instead and manually remove the
            // place holders
            PowerPointSlide newSlide = AddSlide(name: name);
            newSlide.DeleteAllShapes();

            Shape categoryNameBox;

            if (fromClipBoard)
            {
                if (!PPLClipboard.Instance.IsLocked)
                {
                    throw new Exception("Clipboard is not locked before pasting!");
                }
                newSlide.Shapes.Paste();
                categoryNameBox = RetrieveCategoryNameBox(newSlide);
            }
            else
            {
                categoryNameBox = newSlide.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0,
                                                             SlideWidth, 0);
                categoryNameBox.TextFrame.TextRange.Text = string.Format(CategoryNameFormat, name);
            }

            Categories.Add(name);
            _categoryNameBoxCollection.Add(categoryNameBox);

            if (setAsDefault)
            {
                _defaultCategory = newSlide;
            }

            Save();
            FlushUndoHistory();
        }
        public string AddShape(PowerPointPresentation pres, PowerPointSlide origSlide, ShapeRange shapeRange, string name, string category = "", bool fromClipBoard = false)
        {
            if (!fromClipBoard)
            {
                return ClipboardUtil.RestoreClipboardAfterAction(() =>
                {
                    return PPLClipboard.Instance.LockAndRelease(() =>
                    {
                        shapeRange.Copy();
                        return AddShape(name, category);
                    });
                }, pres, origSlide);
            }
            else
            {
                if (!PPLClipboard.Instance.IsLocked)
                {
                    throw new Exception("Clipboard is not locked before copying!");
                }

                return AddShape(name, category);
            }
        }

        public override void Close()
        {
            base.Close();

            RetrieveShapeGalleryFile();
        }

        public void CopyCategory(string name)
        {
            if (!PPLClipboard.Instance.IsLocked)
            {
                throw new Exception("Clipboard is not locked before copying!");
            }
            int index = FindCategoryIndex(name);
            Presentation.Slides[index].Shapes.Range().Copy();
        }

        public void CopyShape()
        {
            if (!PPLClipboard.Instance.IsLocked)
            {
                throw new Exception("Clipboard is not locked before copying!");
            }
            _defaultCategory.Shapes.Range().Copy();
        }

        public void CopyShape(string name)
        {
            if (!PPLClipboard.Instance.IsLocked)
            {
                throw new Exception("Clipboard is not locked before copying!");
            }
            List<Shape> shapes = _defaultCategory.GetShapesWithRule(GenerateNameSearchPattern(name));

            _defaultCategory.Shapes.Range(shapes.Select(item => item.Name).ToArray()).Copy();
        }

        public void CopyShape(IEnumerable<string> nameList)
        {
            if (!PPLClipboard.Instance.IsLocked)
            {
                throw new Exception("Clipboard is not locked before copying!");
            }
            List<string> fullList = new List<string>();

            foreach (string name in nameList)
            {
                fullList.AddRange(_defaultCategory.GetShapesWithRule(GenerateNameSearchPattern(name))
                                                  .Select(item => item.Name));
            }

            _defaultCategory.Shapes.Range(fullList.ToArray()).Copy();
        }

        public void CopyShapeToCategory(PowerPointPresentation pres, PowerPointSlide origSlide, string name, string categoryName)
        {
            int index = FindCategoryIndex(categoryName);

            if (index == -1)
            {
                return;
            }

            // copy a shape with name from default category to another category
            List<Shape> shapes = _defaultCategory.GetShapesWithRule(GenerateNameSearchPattern(name));
            PowerPointSlide destCategory = Slides[index - 1];

            ClipboardUtil.RestoreClipboardAfterAction(() =>
            {
                ShapeRange result = PPLClipboard.Instance.LockAndRelease(() =>
                {
                    _defaultCategory.Shapes.Range(shapes.Select(item => item.Name).ToArray()).Copy();
                    return destCategory.Shapes.Paste();
                });
                return result;
            }, pres, origSlide);


            Save();
            FlushUndoHistory();
        }

        public bool HasCategory(string name)
        {
            return Slides.Any(category => category.Name == name);
        }

        public void MoveShapeToCategory(PowerPointPresentation pres, PowerPointSlide origSlide, string name, string categoryName)
        {
            int index = FindCategoryIndex(categoryName);

            if (index == -1)
            {
                return;
            }

            // move a shape with name from default category to another category
            List<Shape> shapes = _defaultCategory.GetShapesWithRule(GenerateNameSearchPattern(name));
            PowerPointSlide destCategory = Slides[index - 1];

            ClipboardUtil.RestoreClipboardAfterAction(() =>
            {
                ShapeRange result = PPLClipboard.Instance.LockAndRelease(() =>
                {
                    _defaultCategory.Shapes.Range(shapes.Select(item => item.Name).ToArray()).Cut();
                    return destCategory.Shapes.Paste();
                });
                return result;
            }, pres, origSlide);

            Save();
            FlushUndoHistory();
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

            if (!ConsistencyCheck())
            {
                return false;
            }

            // set default category to be the first slide, but do nothing if the presentation
            // has no slide, i.e. it's a newly created presentation
            if (Presentation.Slides.Count > 0)
            {
                _defaultCategory = PowerPointSlide.FromSlideFactory(Presentation.Slides[1]);  
            }

            return true;
        }

        public void RemoveCategory()
        {
            // we need to change the index to 0-based in order to remove from Categories
            int index = FindCategoryIndex(_defaultCategory.Name) - 1;

            _defaultCategory = null;
            
            Categories.RemoveAt(index);
            _categoryNameBoxCollection.RemoveAt(index);

            RemoveSlide(index);

            Save();
            FlushUndoHistory();
        }

        public void RemoveShape(string name)
        {
            _defaultCategory.DeleteShapeWithRule(GenerateNameSearchPattern(name));
            
            Save();
            FlushUndoHistory();
        }

        public void RenameShape(string oldName, string newName)
        {
            Regex nameRegex = GenerateNameSearchPattern(oldName);
            Regex replaceRegex = new Regex(oldName);
            List<Shape> shapes = _defaultCategory.GetShapesWithRule(nameRegex);

            foreach (Shape shape in shapes)
            {
                shape.Name = replaceRegex.Replace(shape.Name, newName);
            }

            Save();
            FlushUndoHistory();
        }

        public void RenameCategory(string newName)
        {
            Categories[_defaultCategory.Index - 1] = newName;
            _defaultCategory.Name = newName;

            Shape categoryNameBox = _categoryNameBoxCollection[_defaultCategory.Index - 1];
            categoryNameBox.TextFrame.TextRange.Text = string.Format(CategoryNameFormat, newName);

            Save();
            FlushUndoHistory();
        }
        # endregion

        # region Helper Function

        /// <summary>
        /// Flushes the undo history with a dummy action.
        /// </summary>
        private void FlushUndoHistory()
        {
            for (int i = 0; i < MaxUndoAmount; i++)
            {
                Presentation.Slides[1].Background.Fill.BackColor = Presentation.Slides[1].Background.Fill.BackColor;
            }
        }

        private bool ConsistencyCheck()
        {
            // if the opening ShapeGallery is a single shape file, or if there's no slide,
            // the file is always valid
            return (IsImportedFile && !string.IsNullOrEmpty(ImportToCategory)) ||
                   SlideCount < 1 ||
                   InitCategories();
        }

        private Shape ConsistencyCheckCategoryNameBox(PowerPointSlide category, ref int untitledCategoryCnt)
        {
            Shape categoryNameBox = RetrieveCategoryNameBox(category);

            if (categoryNameBox != null)
            {
                category.Name = RetrieveCategoryName(categoryNameBox);
            }
            else
            {
                // if we do not have a name box inside, we have 3 cases:
                // 1. slide.Name has been configured (old ShapeGallery file);
                // 2. slide.Name is default (user didn't specify a name).

                // for case 1 & 2, we need to add a new text box into the slie.
                // For case 1, the text of category box should be slide.Name;
                // For case 2, the text of category box should be next untitled name;
                categoryNameBox = category.Shapes.AddTextbox(Office.MsoTextOrientation.msoTextOrientationHorizontal, 0, 0,
                                                             SlideWidth, 0);

                Regex defaultSlideNameRegex = new Regex(DefaultSlideNameSearchPattern);

                if (defaultSlideNameRegex.IsMatch(category.Name))
                {
                    untitledCategoryCnt++;
                    
                    string untitledName = string.Format(UntitledCategoryNameFormat, untitledCategoryCnt);
                    category.Name = untitledName;
                }

                categoryNameBox.TextFrame.TextRange.Text = string.Format(CategoryNameFormat, category.Name);
            }

            _categoryNameBoxCollection.Add(categoryNameBox);
            
            return categoryNameBox;
        }

        private bool ConsistencyCheckCategoryLocalToSlide()
        {
            List<string> categoriesOnDisk = Directory.EnumerateDirectories(Path).ToList();
            bool categoryLost = false;

            foreach (string categoryPath in categoriesOnDisk)
            {
                string categoryName = new DirectoryInfo(categoryPath).Name;

                if (Slides.All(category => category.Name.ToLower() != categoryName.ToLower()))
                {
                    categoryLost = true;
                    AddCategory(categoryName, false, false);
                }
            }

            return categoryLost;
        }

        private string ConsistencyCheckCategorySlideToLocal(PowerPointSlide category)
        {
            string categoryFolderPath = System.IO.Path.Combine(Path, category.Name);
            string newCategoryPath = categoryFolderPath;

            // the category is some how lost on the disk, regenerate the category
            if (!Directory.Exists(categoryFolderPath))
            {
                // create the directory, since shape reconstruction will be taken care
                // of during ConsistencyCheckShapeToPng(), we do not need to generate
                // the shapes here
                Directory.CreateDirectory(categoryFolderPath);
            }
            else
            {
                // in case some of categories to be imported have the same name as those
                // already exist categories
                if (IsImportedFile)
                {
                    int duplicateCnt = 1;
                    string oriCategoryName = newCategoryPath;

                    while (Directory.Exists(newCategoryPath))
                    {
                        newCategoryPath = oriCategoryName + " " + duplicateCnt;
                        duplicateCnt++;
                    }

                    Directory.CreateDirectory(newCategoryPath);
                }
            }

            return newCategoryPath;
        }

        private bool ConsistencyCheckPngToShape(IEnumerable<string> pngShapes, PowerPointSlide category)
        {
            // if some png could not be found in shape gallery, we will delete it
            // to save space
            bool shapeLost = false;

            foreach (string pngShape in pngShapes)
            {
                string shapeName = System.IO.Path.GetFileNameWithoutExtension(pngShape);
                Regex searchPattern = GenerateNameSearchPattern(shapeName);
                bool found = category.HasShapeWithRule(searchPattern);

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
            bool shapeDuplicate = false;

            // we have 3 cases here:
            // 1. Open ShapeGallery;
            // 2. Open a ShapeGallery via ImportCategory;

            // For both cases, if inconsistency is found, we keep all the shapes but
            // append "(recovered shape X)" to the shape name, X is the relative index
            // Note: point 2 is not needed, becuase all no-png shapes will be exported
            // during ConsistencyCheckShapeToPng, and pngs without a corresponding shape
            // will be deleted during ConsistencyCheckPngToShape.
            foreach (PowerPointSlide category in Slides)
            {
                Dictionary<string, int> shapeHash = new Dictionary<string, int>();
                List<Shape> shapes = category.Shapes.Cast<Shape>().ToList();
                List<string> duplicateShapeNames = new List<string>();

                foreach (Shape shape in shapes)
                {
                    if (shapeHash.Count == 0 ||
                        !shapeHash.ContainsKey(shape.Name))
                    {
                        shapeHash[shape.Name] = 1;
                    }
                    else
                    {
                        int index = (shapeHash[shape.Name] += 1);

                        // add to collection only if this shape is the first duplicate shape
                        if (index == 2)
                        {
                            duplicateShapeNames.Add(shape.Name);
                        }

                        shape.Name += string.Format(DuplicateShapeSuffixFormat, index);
                    }
                }

                shapeDuplicate = duplicateShapeNames.Count > 0;

                foreach (string lastShapeName in duplicateShapeNames)
                {
                    Shape lastShape = category.GetShapeWithName(lastShapeName)[0];

                    lastShape.Name += string.Format(DuplicateShapeSuffixFormat, 1);
                }
            }

            return shapeDuplicate;
        }

        private bool ConsistencyCheckShapeToPng(List<string> pngShapes, PowerPointSlide category, string shapeFolderPath)
        {
            // if inconsistency is found, we export the extra shape to .png
            bool shapeLost = false;
            Regex groupSelectNamePattern = new Regex(GroupSelectionNamePattern);

            // this is to handle 2 cases:
            // 1. user deleted the .png shape accidentally;
            // 2. the file is imported
            foreach (Shape shape in category.Shapes)
            {
                // skip category name box
                if (shape.Type == Office.MsoShapeType.msoTextBox &&
                    _categoryNameBoxCollection.Contains(shape))
                {
                    continue;
                }

                string name = shape.Name;

                //check for sequence grouped shape
                if (groupSelectNamePattern.IsMatch(name))
                {
                    name = groupSelectNamePattern.Match(name).Groups[1].Value;
                }

                string shapePath = shapeFolderPath + @"\" + name + ".png";

                if (!pngShapes.Contains(shapePath))
                {
                    GraphicsUtil.ExportShape(shape, shapePath);
                    shapeLost = true;
                }
            }

            return shapeLost;
        }

        private int FindCategoryIndex(string categoryName, bool setAsDefault = false)
        {
            int index = -1;

            foreach (PowerPointSlide category in Slides)
            {
                if (category.Name == categoryName)
                {
                    index = category.Index;

                    if (setAsDefault)
                    {
                        _defaultCategory = category;
                    }
                }
            }

            return index;
        }

        private Regex GenerateNameSearchPattern(string name)
        {
            string skippedName = CommonUtil.SkipRegexCharacter(name);
            string searchPattern = string.Format(NameSearchPattern, skippedName, skippedName);
            return new Regex(searchPattern);
        }

        private bool InitCategories()
        {
            // here we need to check 3 cases:
            // 1. self consistency check (if there are any duplicate names);
            // 2. more png than shapes inside pptx (shapes for short);
            // 3. more shapes than png.

            bool shapeDuplicate = ConsistencyCheckSelf();
            bool shapeLost = false;
            bool pngLost = false;
            int untitledCategoryCnt = 0;
            List<PowerPointSlide> slides = Slides;

            for (int i = Slides.Count - 1; i >= 0; i--)
            {
                PowerPointSlide category = Slides[i];
                Shape categoryNameBox = ConsistencyCheckCategoryNameBox(category, ref untitledCategoryCnt);

                // check if we have a corresponding category directory in the Path
                string shapeFolderPath;
                try
                {
                    shapeFolderPath = ConsistencyCheckCategorySlideToLocal(category);
                }
                catch (Exception)
                {
                    // Unable to get shape folder path. Store the problematic category and continue
                    // Actual slide index starts from 1, but is accounted for in PowerPointPresentation.removeSlide(int).
                    RemoveSlide(i);
                    continue;
                }
                string finalCategoryName = new DirectoryInfo(shapeFolderPath).Name;

                List<string> pngShapes = Directory.EnumerateFiles(shapeFolderPath, "*.png").ToList();

                // critical: OR with itself at the end to avoid early termination
                shapeLost = ConsistencyCheckShapeToPng(pngShapes, category, shapeFolderPath) || shapeLost;
                pngLost = ConsistencyCheckPngToShape(pngShapes, category) || pngLost;

                // update names only when the name gets changed
                if (category.Name != finalCategoryName)
                {
                    category.Name = finalCategoryName;
                    categoryNameBox.TextFrame.TextRange.Text = string.Format(CategoryNameFormat, finalCategoryName);
                }

                Categories.Add(finalCategoryName);
            }

            bool categoryInShapeGalleryLost = ConsistencyCheckCategoryLocalToSlide();

            Save();

            if ((shapeDuplicate || shapeLost || categoryInShapeGalleryLost || pngLost) &&
                !IsImportedFile)
            {
                MessageBox.Show(ShapesLabText.ErrorShapeCorrupted);

                return false;
            }

            return true;
        }

        private string RetrieveCategoryName(Shape categoryNameBox)
        {
            Regex categoryNamePattern = new Regex(CategoryNameBoxSearchPattern);
            Match namePatternMatch = categoryNamePattern.Match(categoryNameBox.TextFrame.TextRange.Text);
            string categoryName = namePatternMatch.Groups[1].Value;

            return categoryName;
        }

        private Shape RetrieveCategoryNameBox(PowerPointSlide slide)
        {
            List<Shape> nameBoxCandidate = slide.GetShapesWithTypeAndRule(Office.MsoShapeType.msoTextBox, new Regex(".+"));

            if (nameBoxCandidate.Count == 0)
            {
                return null;
            }

            Regex categoryNamePattern = new Regex(CategoryNameBoxSearchPattern);

            // return the first match name box
            return nameBoxCandidate.FirstOrDefault(x => categoryNamePattern.IsMatch(x.TextFrame.TextRange.Text));
        }

        private void RetrievePptxFile()
        {
            string shapeGalleryFileName = FullName.Replace(".pptx", ShapeGalleryFileExtension);

            if (File.Exists(shapeGalleryFileName))
            {
                File.SetAttributes(shapeGalleryFileName, FileAttributes.Normal);
                // To ensure that the file fullName does not exist before attempting to move
                if (!File.Exists(FullName))
                {
                    File.Move(shapeGalleryFileName, FullName);
                }
            }

            if (File.Exists(FullName))
            {
                // to reduce the chance that user opens the shape gallery file, we make the pptx file hidden
                File.SetAttributes(FullName, FileAttributes.Normal);
                File.SetAttributes(FullName, FileAttributes.Hidden);
            }
        }

        private void RetrieveShapeGalleryFile()
        {
            // set the file as a visible readonly .pptlabsshapes file.
            string shapeGalleryFileName = FullName.Replace(".pptx", ShapeGalleryFileExtension);

            Trace.TraceInformation("FullName = " + FullName + ", Name = " + shapeGalleryFileName);

            File.Move(FullName, shapeGalleryFileName);

            File.SetAttributes(shapeGalleryFileName, FileAttributes.Normal);
            File.SetAttributes(shapeGalleryFileName, FileAttributes.ReadOnly);
        }

        private string AddShape(string name, string category) 
        {
            PowerPointSlide categorySlide = _defaultCategory;

            if (!string.IsNullOrEmpty(category))
            {
                int categoryIndex = FindCategoryIndex(category);

                if (categoryIndex == -1)
                {
                    return string.Empty;
                }

                categorySlide = Slides[categoryIndex - 1];
            }

            // check if the name has been used, if used, name it to the next available name
            if (categorySlide.HasShapeWithRule(GenerateNameSearchPattern(name)))
            {
                Regex nameExtractionRegex = new Regex(string.Format(NameExtractionPatternFormat, name, name));
                List<string> nameList = categorySlide.GetShapesWithRule(nameExtractionRegex)
                                            .Select(item => nameExtractionRegex.Match(item.Name))
                                            .Select(match => !string.IsNullOrEmpty(match.Groups[1].Value)
                                                             ? match.Groups[1].Value
                                                             : match.Groups[2].Value)
                                            .Distinct()
                                            .ToList();

                name = CommonUtil.NextAvailableName(nameList, name);
            }

            ShapeRange pastedShapeRange = categorySlide.Shapes.Paste();

            if (pastedShapeRange.Count > 1)
            {
                for (int nameCount = 1; nameCount <= pastedShapeRange.Count; nameCount++)
                {
                    Shape shape = pastedShapeRange[nameCount];

                    shape.Name = string.Format(GroupSelectionNameFormat, name, nameCount);
                }
            }
            else
            {
                pastedShapeRange[1].Name = name;
            }

            Save();
            FlushUndoHistory();

            return name;
        }
        # endregion
    }
}
