using System.ComponentModel;
using PowerPointLabs.Annotations;

namespace PowerPointLabs.ImageSearch.Model
{
    public class ImageItem : INotifyPropertyChanged
    {
        // thumbnail-purpose
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

        public string BlurImageFile { get; set; }

        public string FullSizeImageFile { get; set; }

        private string _fullSizeImageUri;

        public string FullSizeImageUri
        {
            get
            {
                return _fullSizeImageUri;
                
            }
            set
            {
                _fullSizeImageUri = value;
                OnPropertyChanged("FullSizeImageUri");
            }
        }

        public event PropertyChangedEventHandler PropertyChanged;

        [NotifyPropertyChangedInvocator]
        protected virtual void OnPropertyChanged(string propertyName)
        {
            var handler = PropertyChanged;
            if (handler != null) handler(this, new PropertyChangedEventArgs(propertyName));
        }
    }
}
