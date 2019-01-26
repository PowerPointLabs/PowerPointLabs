using System.Windows;

using PowerPointLabs.PictureSlidesLab.Util;

namespace PowerPointLabs.PictureSlidesLab.Model
{
    public class ImageItem : WPF.Observable.Model
    {
        # region UI related prop
        private string _imageFile;

        public string ImageFile
        {
            get { return _imageFile; }
            set
            {
                _imageFile = value;
                OnPropertyChanged("ImageFile");
            }
        }

        private string _tooltip;

        public string Tooltip
        {
            get
            {
                return _tooltip;
            }
            set
            {
                _tooltip = value;
                OnPropertyChanged("Tooltip");
            }
        }
        # endregion

        # region other info
        // as cache
        public string BlurImageFile { get; set; }
        public string SpecialEffectImageFile { get; set; }
        public string FullSizeImageFile { get; set; }

        // picture dimensions adjustment related
        public string CroppedImageFile { get; set; }
        public string CroppedThumbnailImageFile { get; set; }
        // define picture dimensions (cropped region)
        public Rect Rect { get; set; }

        // meta info
        public string ContextLink { get; set; }
        public string Source { get; set; }

        // backup info
        public string BackupFullSizeImageFile { get; set; }
        #endregion

        public void UpdateDownloadedImage(string imagePath)
        {
            FullSizeImageFile = imagePath;
            ImageFile = ImageUtil.GetThumbnailFromFullSizeImg(FullSizeImageFile);
            Tooltip = ImageUtil.GetWidthAndHeight(FullSizeImageFile);
        }

        public void UpdateImageAdjustmentOffset(string adjustResult, string thumbnail, Rect rect)
        {
            CroppedImageFile = adjustResult;
            CroppedThumbnailImageFile = thumbnail;
            Rect = rect;
        }
    }
}
