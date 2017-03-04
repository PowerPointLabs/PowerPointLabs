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
        private const string PictureSlidesLabImagesList = "PictureSlidesLabImagesList";
        private const string PictureSlidesLabSettings = "PictureSlidesLabSettings";
        private const string PictureSlidesLabCustomStyles = "PictureSlidesLabCustomStyles";

        public static string AggregatedFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\" + "pptlabs_pictureSlidesLab" + @"\";

        public static readonly string LoadingImgPath = AggregatedFolder + "loading";
        public static readonly string ChoosePicturesImgPath = AggregatedFolder + "choosePicture";
        public static readonly string NoPicturePlaceholderImgPath = AggregatedFolder + "noPicturePlaceholder";
        public static readonly string SampleImg1Path = AggregatedFolder + "sample1";
        public static readonly string SampleImg2Path = AggregatedFolder + "sample2";

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
                ErrorDialogWrapper.ShowDialog("Failed to remove unused images.", e.Message, e);
            }
        }

        private static void Empty(DirectoryInfo directory, ICollection<string> filesInUse)
        {
            if (!directory.Exists) return;

            try
            {
                filesInUse.Add(AggregatedFolder + PictureSlidesLabImagesList);
                filesInUse.Add(AggregatedFolder + PictureSlidesLabSettings);
                filesInUse.Add(AggregatedFolder + PictureSlidesLabCustomStyles);
                filesInUse.Add(LoadingImgPath);
                filesInUse.Add(ChoosePicturesImgPath);
                filesInUse.Add(NoPicturePlaceholderImgPath);
                filesInUse.Add(SampleImg1Path);
                filesInUse.Add(SampleImg2Path);
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
                Resources.ChoosePicturesIcon.Save(ChoosePicturesImgPath);
                Resources.DefaultPicture.Save(NoPicturePlaceholderImgPath);
                Resources.PslSample1.Save(SampleImg1Path);
                Resources.PslSample2.Save(SampleImg2Path);
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
                using (var writer = new StreamWriter(GetPath(PictureSlidesLabSettings)))
                {
                    var serializer = new XmlSerializer(settings.GetType());
                    serializer.Serialize(writer, settings);
                    writer.Flush();
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Failed to save Picture Slides Lab settings");
            }
        }

        /// Taken from http://stackoverflow.com/a/14663848
        /// <summary>
        /// Saves style to an xml file
        /// </summary>
        /// <param name="filename">File path of the new xml file</param>
        public static void Save(StyleOption styleOption, string filename)
        {
            try
            {
                using (var writer = new StreamWriter(filename))
                {
                    var serializer = new XmlSerializer(styleOption.GetType());
                    serializer.Serialize(writer, styleOption);
                    writer.Flush();
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Failed to save Picture Slides Lab Style Options: " + e.StackTrace);
            }
        }
        
        /// <summary>
        /// Save user-customized style list
        /// </summary>
        /// <param name="styleOption"></param>
        public static void Save(List<StyleOption> styleOptions)
        {
            try
            {
                using (var writer = new StreamWriter(GetPath(PictureSlidesLabCustomStyles)))
                {
                    var serializer = new XmlSerializer(styleOptions.GetType());
                    serializer.Serialize(writer, styleOptions);
                    writer.Flush();
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Failed to save Picture Slides Lab custom styles");
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
                using (var stream = File.OpenRead(GetPath(PictureSlidesLabSettings)))
                {
                    var serializer = new XmlSerializer(typeof(Model.Settings));
                    var settings = serializer.Deserialize(stream) as Model.Settings
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
        
        /// <summary>
        /// Load an object from an xml file
        /// </summary>
        /// <param name="filename">Xml file name</param>
        /// <returns>The object created from the xml file</returns>
        public static StyleOption LoadStyleOption(string filename)
        {
            try
            {
                using (var stream = File.OpenRead(filename))
                {
                    var serializer = new XmlSerializer(typeof(StyleOption));
                    var opt = serializer.Deserialize(stream) as StyleOption;
                    return opt ?? new StyleOption();
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Failed to load Picture Slides Lab Style Options: " + e.StackTrace);
                return new StyleOption();
            }
        }

        /// <summary>
        /// Load user-customized style list
        /// </summary>
        /// <returns></returns>
        public static List<StyleOption> LoadCustomStyles()
        {
            try
            {
                using (var stream = File.OpenRead(GetPath(PictureSlidesLabCustomStyles)))
                {
                    var serializer = new XmlSerializer(typeof(List<StyleOption>));
                    var styleOptions = serializer.Deserialize(stream) as List<StyleOption>
                        ?? new List<StyleOption>();
                    return styleOptions;
                }
            }
            catch (Exception e)
            {
                Logger.LogException(e, "Failed to load Picture Slides Lab custom styles");
                return new List<StyleOption>();
            }
        }
    }
}
