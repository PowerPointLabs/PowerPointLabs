using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using NAudio.Wave;
using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.ELearningLab.Utility
{
    public class StreamUtility
    {
        public static void SaveStreamToFile(string fileFullPath, Stream stream)
        {
            try
            {
                byte[] bytesInStream = ReadFully(stream);
                using (WaveFileWriter writer = new WaveFileWriter(fileFullPath, new WaveFormat(22050, 1)))
                {
                    writer.Write(bytesInStream, 0, bytesInStream.Length);
                }
            }
            catch
            {
                Logger.Log("Failed to save stream to .wav file");
            }
        }
        private static byte[] ReadFully(Stream input)
        {
            byte[] buffer = new byte[16 * 1024];

            using (MemoryStream ms = new MemoryStream())
            {
                int read;
                while ((read = input.Read(buffer, 0, buffer.Length)) > 0)
                {
                    ms.Write(buffer, 0, read);
                }
                return ms.ToArray();
            }
        }
    }
}
