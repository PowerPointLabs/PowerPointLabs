using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.Xml.Serialization;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.Properties;
using PowerPointLabs.Views;

namespace PowerPointLabs.ImageSearch.Util
{
    class StoragePath
    {
        public static string AggregatedFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\" + "pptlabs_imagesLab" + @"\";

        public static readonly string LoadingImgPath = AggregatedFolder + "loading";
        public static readonly string LoadMoreImgPath = AggregatedFolder + "loadMore";
        public static readonly string QuickDropDialogSettingsPath = AggregatedFolder + "quick-drop-dialog.xml";

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
            return true;
        }

        private static void Empty(DirectoryInfo directory, ICollection<string> filesInUse)
        {
            try
            {
                filesInUse.Add(AggregatedFolder + "ImagesLabImagesList");
                filesInUse.Add(LoadingImgPath);
                filesInUse.Add(LoadMoreImgPath);
                filesInUse.Add(QuickDropDialogSettingsPath);
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
                Resources.LoadMore.Save(LoadMoreImgPath);
            }
            catch
            {
                // may fail to save it, which is fine
            }
        }

        public static string GetPath(string name)
        {
            return AggregatedFolder + name;
        }

        /// <summary>
        /// Save window info (window positions etc)
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="windowInfo"></param>
        public static void Save(string filename, WindowInfo windowInfo)
        {
            try
            {
                using (var writer = new StreamWriter(GetPath(filename)))
                {
                    var serializer = new XmlSerializer(windowInfo.GetType());
                    serializer.Serialize(writer, windowInfo);
                    writer.Flush();
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.Log("Failed to save Images Lab window info: " + e.StackTrace, "Error");
            }
        }

        /// <summary>
        /// Save images list
        /// </summary>
        /// <param name="filename"></param>
        /// <param name="list"></param>
        public static void Save(string filename, Collection<ImageItem> list)
        {
            try
            {
                using (var writer = new StreamWriter(GetPath(filename)))
                {
                    var serializer = new XmlSerializer(list.GetType());
                    serializer.Serialize(writer, list);
                    writer.Flush();
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.Log("Failed to save Images Lab settings: " + e.StackTrace, "Error");
            }
        }

        /// <summary>
        /// Load window info
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static WindowInfo LoadWindowInfo(string filename)
        {
            try
            {
                using (var stream = File.OpenRead(GetPath(filename)))
                {
                    var serializer = new XmlSerializer(typeof(WindowInfo));
                    var result = serializer.Deserialize(stream) as WindowInfo ?? new WindowInfo();
                    return result;
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.Log("Failed to load Images Lab window info: " + e.StackTrace, "Error");
                return new WindowInfo();
            }
        }

        /// <summary>
        /// Load images list
        /// </summary>
        /// <param name="filename"></param>
        /// <returns></returns>
        public static List<ImageItem> Load(string filename)
        {
            try
            {
                using (var stream = File.OpenRead(GetPath(filename)))
                {
                    var serializer = new XmlSerializer(typeof(List<ImageItem>));
                    var list = serializer.Deserialize(stream) as List<ImageItem> ?? new List<ImageItem>();
                    return list;
                }
            }
            catch (Exception e)
            {
                PowerPointLabsGlobals.Log("Failed to load Images Lab settings: " + e.StackTrace, "Error");
                return new List<ImageItem>();
            }
        }
    }
}
