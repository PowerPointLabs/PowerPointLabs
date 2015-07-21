
namespace PowerPointLabs.ImageSearch.Model
{
    public class ImageItem : Notifiable
    {
        // to be shown in the UI
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

        // as cache
        public string BlurImageFile { get; set; }

        // as cache
        public string FullSizeImageFile { get; set; }

        // meta info
        public bool IsToDelete { get; set; }

        private string _fullSizeImageUri;

        public string FullSizeImageUri
        {
            get { return _fullSizeImageUri; }
            set
            {
                _fullSizeImageUri = value;
                OnPropertyChanged("FullSizeImageUri");
            }
        }
    }
}
