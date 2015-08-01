namespace PowerPointLabs.ImageSearch.Domain
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

        // meta info
        public bool IsToDelete { get; set; }
        public string ContextLink { get; set; }
        public string FullSizeImageUri { get; set; }
        # endregion
    }
}
