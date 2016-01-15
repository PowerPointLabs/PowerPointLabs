using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Xml.Serialization;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.Properties;
using PowerPointLabs.Views;

namespace PowerPointLabs.PictureSlidesLab.Util
{
    public class StoragePath
    {
        private const string PictureSlidesLabImagesList = "PictureSlidesLabImagesList";

        public static string AggregatedFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\" + "pptlabs_pictureSlidesLab" + @"\";

        public static readonly string LoadingImgPath = AggregatedFolder + "loading";

        private static bool _isInit;

        public static bool InitPersistentFolder(ICollection<string> filesInUse)
        {
            try
            {
                Empty(new DirectoryInfo(AggregatedFolder), filesInUse);
            }
            catch (Exception e)
            {
                ErrorDialogWrapper.ShowDialog("Failed to remove unused images.", e.Message, e);
            }

            if (!Directory.Exists(AggregatedFolder))
            {
                try
                {
                    Directory.CreateDirectory(AggregatedFolder);
                }
                catch
                {
                    return false;
                }
            }

            InitResources();
            _isInit = true;
            return true;
        }

        private static void Empty(DirectoryInfo directory, ICollection<string> filesInUse)
        {
            if (!directory.Exists) return;

            try
            {
                filesInUse.Add(AggregatedFolder + PictureSlidesLabImagesList);
                filesInUse.Add(LoadingImgPath);
                foreach (var file in directory.GetFiles())
                {
                    if (!filesInUse.Contains(file.FullName))
                    {
                        try
                        {
                            file.Delete();
                        }
                        catch
                        {
                            // may be still in use, which is fine
                        }
                    }
                }
            }
            catch (Exception)
            {
                // ignore ex, if cannot delete trash
            }
        }

        private static void InitResources()
        {
            try
            {
                Resources.Loading.Save(LoadingImgPath);
            }
            catch
            {
                // may fail to save it, which is fine
            }
        }

        public static string GetPath(string name)
        {
            if (!_isInit)
            {
                throw new Exception("StoragePath is not initialized!");
            }
            return AggregatedFolder + name;
        }

        /// <summary>
        /// Save images list
        /// </summary>
        /// <param name="list"></param>
        public static void Save(Collection<ImageItem> list)
        {
            try
            {
                using (var writer = new StreamWriter(GetPath(PictureSlidesLabImagesList)))
                {
                    var serializer = new XmlSerializer(list.GetType());
                    serializer.Serialize(writer, list);
                    writer.Flush();
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.Log("Failed to save Picture Slides Lab settings: " + e.StackTrace, "Error");
            }
        }

        /// <summary>
        /// Load images list
        /// </summary>
        /// <returns></returns>
        public static ObservableCollection<ImageItem> Load()
        {
            try
            {
                using (var stream = File.OpenRead(GetPath(PictureSlidesLabImagesList)))
                {
                    var serializer = new XmlSerializer(typeof(ObservableCollection<ImageItem>));
                    var list = serializer.Deserialize(stream) as ObservableCollection<ImageItem> 
                        ?? new ObservableCollection<ImageItem>();
                    return list;
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.Log("Failed to load Picture Slides Lab settings: " + e.StackTrace, "Error");
                return new ObservableCollection<ImageItem>();
            }
        }
    }
}
