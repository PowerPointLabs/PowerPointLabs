using System.ComponentModel;

using ImageProcessor.Imaging.Filters;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    partial class StyleOption
    {
        #region APIs
        public IMatrixFilter GetSpecialEffect()
        {
            switch (SpecialEffect)
            {
                case 0:
                    return MatrixFilters.GreyScale;
                case 1:
                    return MatrixFilters.BlackWhite;
                case 2:
                    return MatrixFilters.Comic;
                case 3:
                    return MatrixFilters.Gotham;
                case 4:
                    return MatrixFilters.HiSatch;
                case 5:
                    return MatrixFilters.Invert;
                case 6:
                    return MatrixFilters.Lomograph;
                case 7:
                    return MatrixFilters.LoSatch;
                case 8:
                    return MatrixFilters.Polaroid;
                // case 9:
                default:
                    return MatrixFilters.Sepia;
            }
        }
        #endregion

        [DefaultValue(false)]
        public bool IsUseSpecialEffectStyle { get; set; }

        [DefaultValue(-1)]
        public int SpecialEffect { get; set; }
    }
}
