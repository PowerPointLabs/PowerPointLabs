using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Xml.Serialization;

using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.PictureSlidesLab.Model;
using PowerPointLabs.Properties;
using PowerPointLabs.Views;

namespace PowerPointLabs.PictureSlidesLab.Util
{
    public class StoragePath
    {
        public static string AggregatedFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\" + "pptlabs_pictureSlidesLab" + @"\";

        public static readonly string LoadingImgPath = AggregatedFolder + "loading";
        public static readonly string ChoosePicturesImgPath = AggregatedFolder + "choosePicture";
        public static readonly string NoPicturePlaceholderImgPath = AggregatedFolder + "noPicturePlaceholder";
        public static readonly string SampleImg1Path = AggregatedFolder + "sample1";
        public static readonly string SampleImg2Path = AggregatedFolder + "sample2";

        private const string PictureSlidesLabImagesList = "PictureSlidesLabImagesList";
        private const string PictureSlidesLabSettings = "PictureSlidesLabSettings";

        private static bool _isInit;
        private static bool _isFirstTimeUsage;

        public static bool InitPersistentFolder()
        {
            if (!Directory.Exists(AggregatedFolder))
            {
                _isFirstTimeUsage = true;
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

        public static bool IsFirstTimeUsage()
        {
            return _isFirstTimeUsage;
        }

        public static void CleanPersistentFolder(ICollection<string> filesInUse)
        {
            try
            {
                Empty(new DirectoryInfo(AggregatedFolder), filesInUse);
            }
            catch (Exception e)
            {
                ErrorDialogBox.ShowDialog("Failed to remove unused images.", e.Message, e);
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
                using (StreamWriter writer = new StreamWriter(GetPath(PictureSlidesLabImagesList)))
                {
                    XmlSerializer serializer = new XmlSerializer(list.GetType());
                    serializer.Serialize(writer, list);
                    writer.Flush();
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Failed to save Picture Slides Lab images list");
            }
        }
        
        /// <summary>
        /// Save PSL settings
        /// </summary>
        /// <param name="settings"></param>
        public static void Save(Model.Settings settings)
        {
            try
            {
                using (StreamWriter writer = new StreamWriter(GetPath(PictureSlidesLabSettings)))
                {
                    XmlSerializer serializer = new XmlSerializer(settings.GetType());
                    serializer.Serialize(writer, settings);
                    writer.Flush();
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Failed to save Picture Slides Lab settings");
            }
        }

        /// <summary>
        /// Load images list
        /// </summary>
        /// <returns></returns>
        public static ObservableCollection<ImageItem> LoadPictures()
        {
            try
            {
                using (FileStream stream = File.OpenRead(GetPath(PictureSlidesLabImagesList)))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(ObservableCollection<ImageItem>));
                    ObservableCollection<ImageItem> list = serializer.Deserialize(stream) as ObservableCollection<ImageItem> 
                        ?? new ObservableCollection<ImageItem>();
                    return list;
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Failed to load Picture Slides Lab images list");
                return new ObservableCollection<ImageItem>();
            }
        }

        /// <summary>
        /// Load PSL settings
        /// </summary>
        /// <returns></returns>
        public static Model.Settings LoadSettings()
        {
            try
            {
                using (FileStream stream = File.OpenRead(GetPath(PictureSlidesLabSettings)))
                {
                    XmlSerializer serializer = new XmlSerializer(typeof(Model.Settings));
                    Model.Settings settings = serializer.Deserialize(stream) as Model.Settings
                        ?? new Model.Settings();
                    return settings;
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Failed to load Picture Slides Lab settings");
                return new Model.Settings();
            }
        }

        private static void Empty(DirectoryInfo directory, ICollection<string> filesInUse)
        {
            if (!directory.Exists)
            {
                return;
            }

            try
            {
                filesInUse.Add(AggregatedFolder + PictureSlidesLabImagesList);
                filesInUse.Add(AggregatedFolder + PictureSlidesLabSettings);
                filesInUse.Add(LoadingImgPath);
                filesInUse.Add(ChoosePicturesImgPath);
                filesInUse.Add(NoPicturePlaceholderImgPath);
                filesInUse.Add(SampleImg1Path);
                filesInUse.Add(SampleImg2Path);
                foreach (FileInfo file in directory.GetFiles())
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
                Properties.Resources.Loading.Save(LoadingImgPath);
                Properties.Resources.ChoosePicturesIcon.Save(ChoosePicturesImgPath);
                Properties.Resources.DefaultPicture.Save(NoPicturePlaceholderImgPath);
                Properties.Resources.PslSample1.Save(SampleImg1Path);
                Properties.Resources.PslSample2.Save(SampleImg2Path);
            }
            catch
            {
                // may fail to save it, which is fine
            }
        }
    }
}
