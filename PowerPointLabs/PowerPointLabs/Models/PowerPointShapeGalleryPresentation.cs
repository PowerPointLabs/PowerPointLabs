using System.Collections.Generic;
using System.IO;
using System.Text.RegularExpressions;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointShapeGalleryPresentation : PowerPointPresentation
    {
        private const string ShapeGalleryFileExtension = ".pptlabsshapes";
        private PowerPointSlide _defaultCategory;

        private readonly Dictionary<string, int> _categoryNameIndexMapper = new Dictionary<string, int>();

        # region Properties
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
            newSlide.DeleteShapeWithRule(new Regex(@"Title \d+"));
            newSlide.DeleteShapeWithRule(new Regex(@"Content Placeholder \d+"));

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

            var shapeGalleryFileName = FullName.Replace(".pptx", ShapeGalleryFileExtension);

            File.Move(FullName, shapeGalleryFileName);
        }

        public void CopyShape(string name)
        {
            var shapes = _defaultCategory.GetShapesWithPrefix(name);

            if (shapes.Count != 1) return;
            
            shapes[0].Copy();
        }

        public override void Open(bool readOnly = false, bool untitled = false,
                                  bool withWindow = true, bool focus = true)
        {
            var shapeGalleryFileName = FullName.Replace(".pptx", ShapeGalleryFileExtension);

            if (File.Exists(shapeGalleryFileName))
            {
                File.Move(shapeGalleryFileName, FullName);
            }

            base.Open(readOnly, untitled, withWindow, focus);

            if (SlideCount > 0)
            {
                foreach (var category in Slides)
                {
                    _categoryNameIndexMapper[category.Name] = category.Index;

                    // this is to handle the case when user deletes the .png image manually but
                    // ShapeGallery.pptx isn't updated
                    var shapeCnt = 1;
                    
                    while (shapeCnt <= category.Shapes.Count)
                    {
                        var shape = category.Shapes[shapeCnt];
                        var shapePath = Path + @"\" + category.Name + @"\" + shape.Name + ".png";

                        if (!File.Exists(shapePath))
                        {
                            shape.Delete();
                        }
                        else
                        {
                            shapeCnt++;
                        }
                    }
                }

                Save();
            }
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
            _defaultCategory.DeleteShapeWithRule(new Regex(name));
        }

        public void RenameShape(string oldName, string newName)
        {
            var shapes = _defaultCategory.GetShapesWithRule(new Regex(oldName));

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
        # endregion
    }
}
