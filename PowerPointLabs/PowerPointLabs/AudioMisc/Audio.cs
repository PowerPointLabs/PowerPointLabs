using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PPExtraEventHelper;

namespace PowerPointLabs.AudioMisc
{
    internal class Audio
    {
        public enum AudioType
        {
            Record,
            Auto
        }

        public const int MM_MCINOTIFY = 0x03B9;
        public const int MCI_NOTIFY_SUCCESS = 0x01;
        public const int MCI_NOTIFY_ABORTED = 0x04;
        public const int MCI_NOTIFY_FAILURE = 0x08;

        private const int MCI_RET_INFO_BUF_LEN = 128;

        private static StringBuilder mciRetInfo;

        public string Name { get; set; }
        public int MatchSciptID { get; set; }
        public string SaveName { get; set; }
        public string Length { get; set; }
        public int LengthMillis { get; set; }
        public AudioType Type { get; set; }

        public static void OpenNewAudio()
        {
            Native.mciSendString("open new type waveaudio alias sound", null, 0, IntPtr.Zero);
        }

        public static void OpenAudio(string name)
        {
            Native.mciSendString("open \"" + name + "\" alias sound", null, 0, IntPtr.Zero);
        }

        public static void CloseAudio()
        {
            Native.mciSendString("close sound", null, 0, IntPtr.Zero);
        }

        public static int GetAudioLength()
        {
            mciRetInfo = new StringBuilder(MCI_RET_INFO_BUF_LEN);
            Native.mciSendString("status sound length", mciRetInfo, MCI_RET_INFO_BUF_LEN, IntPtr.Zero);
            return Int32.Parse(mciRetInfo.ToString());
        }

        public static int GetAudioLength(string name)
        {
            OpenAudio(name);
            int length = GetAudioLength();
            CloseAudio();

            return length;
        }

        public static int GetAudioCurrentPosition()
        {
            mciRetInfo = new StringBuilder(MCI_RET_INFO_BUF_LEN);
            Native.mciSendString("status sound position", mciRetInfo, MCI_RET_INFO_BUF_LEN, IntPtr.Zero);
            return Int32.Parse(mciRetInfo.ToString());
        }

        public static Audio.AudioType GetAudioType(string name)
        {
            if (name.Contains("Rec"))
            {
                return Audio.AudioType.Record;
            }

            return Audio.AudioType.Auto;
        }
    }
}
