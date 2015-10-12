using System;
using System.Collections.Generic;
using System.IO;
using System.Xml.Serialization;
using PowerPointLabs.ImageSearch.Domain;
using PowerPointLabs.Properties;

namespace PowerPointLabs.ImageSearch.Util
{
    class StoragePath
    {
        public static string AggregatedFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + @"\" + "pptlabs_imagesLab" + @"\";

        public static readonly string LoadingImgPath = AggregatedFolder + "loading";
        public static readonly string LoadMoreImgPath = AggregatedFolder + "loadMore";

        public static bool InitPersistentFolder()
        {
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

        public static void Save(string filename, List<ImageItem> list)
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
