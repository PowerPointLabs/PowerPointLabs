using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
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
            var newSlide = AddSlide(name: name);

            _categoryNameIndexMapper[name] = Slides.Count;

            if (setAsDefault)
            {
                _defaultCategory = newSlide;
            }
        }

        public void AddShape(Selection selection, string name)
        {
            selection.Copy();

            var pastedShape = _defaultCategory.Shapes.Paste();

            if (pastedShape.Count > 1)
            {
                pastedShape.Group();
            }

            pastedShape.Name = name;
        }

        public void AddShape(Selection selection, string category, string name)
        {
            selection.Copy();

            var categorySlide = Slides[_categoryNameIndexMapper[category]];
            var pastedShape = categorySlide.Shapes.Paste();

            if (pastedShape.Count > 1)
            {
                pastedShape.Group();
            }

            pastedShape.Name = name;
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
