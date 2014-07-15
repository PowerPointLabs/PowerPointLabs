using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;

namespace PowerPointLabs.Models
{
    class PowerPointShapeGalleryPresentation : PowerPointPresentation
    {
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

        public void CopyShape(string name)
        {
            var shapes = _defaultCategory.GetShapesWithPrefix(name);

            if (shapes.Count != 1) return;
            
            shapes[0].Copy();
        }

        public override void Open(bool readOnly = false, bool untitled = false, bool withWindow = true, bool focus = true)
        {
            base.Open(readOnly, untitled, withWindow, focus);

            if (SlideCount > 0)
            {
                foreach (var slide in Slides)
                {
                    _categoryNameIndexMapper[slide.Name] = slide.Index;
                }
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
