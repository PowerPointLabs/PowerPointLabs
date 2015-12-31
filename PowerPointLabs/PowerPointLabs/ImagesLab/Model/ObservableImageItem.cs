namespace PowerPointLabs.ImagesLab.Model
{
    class ObservableImageItem : WPF.Observable.Model
    {
        private ImageItem _imageItem;

        public ImageItem ImageItem
        {
            get { return _imageItem; }
            set
            {
                _imageItem = value;
                OnPropertyChanged("ImageItem");
            }
        }
    }
}
