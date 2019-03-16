using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using Microsoft.VisualStudio.TestTools.UnitTesting;

using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;

namespace Test.UnitTest.ELearningLab.Model
{
    [TestClass]
    public class SelfExplanationClickItemTest
    {
        private SelfExplanationClickItem item;

        [TestInitialize]
        public void Init()
        {
            item = new SelfExplanationClickItem(captionText: "test");
        }

        [TestMethod]
        [TestCategory("UT")]
        public void ClickNoChangedNotification()
        {
            bool notified = false;
            item.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == "ClickNo")
                {
                    notified = true;
                }
            };
            item.ClickNo = 1;
            Assert.IsTrue(notified);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void IsVoiceChangedNotification()
        {
            bool isVoiceNotified = false;
            bool isDummyItemNotified = false;
            item.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == "IsVoice")
                {
                    isVoiceNotified = true;
                }
                if (args.PropertyName == "IsDummyItem")
                {
                    isDummyItemNotified = true;
                }
            };
            item.IsVoice = true;
            Assert.IsTrue(isVoiceNotified);
            Assert.IsTrue(isDummyItemNotified);
        }


        [TestMethod]
        [TestCategory("UT")]
        public void IsCaptionChangedNotification()
        {
            bool isCaptionNotified = false;
            bool isDummyItemNotified = false;
            item.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == "IsCaption")
                {
                    isCaptionNotified = true;
                }
                if (args.PropertyName == "IsDummyItem")
                {
                    isDummyItemNotified = true;
                }
            };
            item.IsCaption = true;
            Assert.IsTrue(isCaptionNotified);
            Assert.IsTrue(isDummyItemNotified);
        }


        [TestMethod]
        [TestCategory("UT")]
        public void IsCalloutChangedNotification()
        {
            bool isCalloutNotified = false;
            bool isDummyItemNotified = false;
            item.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == "IsCallout")
                {
                    isCalloutNotified = true;
                }
                if (args.PropertyName == "IsDummyItem")
                {
                    isDummyItemNotified = true;
                }
            };
            item.IsCallout = true;
            Assert.IsTrue(isCalloutNotified);
            Assert.IsTrue(isDummyItemNotified);
        }


        [TestMethod]
        [TestCategory("UT")]
        public void CalloutTextChangedNotification()
        {
            bool calloutTextChangeNotified = false;
            item.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == "CalloutText")
                {
                    calloutTextChangeNotified = true;
                }
            };
            item.CalloutText = "callout text";
            Assert.IsTrue(calloutTextChangeNotified);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void CaptionTextChangedNotification()
        {
            bool captionTextChangeNotified = false;
            item.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == "CaptionText")
                {
                    captionTextChangeNotified = true;
                }
            };
            item.CaptionText = "callout text";
            item.HasShortVersion = false;
            Assert.IsTrue(captionTextChangeNotified);
            Assert.IsTrue(item.CaptionText.Equals(item.CalloutText));
        }


        [TestMethod]
        [TestCategory("UT")]
        public void VoiceLabelChangedNotification()
        {
            bool voiceLabelChangeNotified = false;
            item.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == "VoiceLabel")
                {
                    voiceLabelChangeNotified = true;
                }
            };
            item.VoiceLabel = "voice label";
            Assert.IsTrue(voiceLabelChangeNotified);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void TriggerIndexChangedNotification()
        {
            bool triggerIndexChangeNotified = false;
            item.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == "TriggerIndex")
                {
                    triggerIndexChangeNotified = true;
                }
            };
            item.TriggerIndex = 1;
            Assert.IsTrue(triggerIndexChangeNotified);
        }


        [TestMethod]
        [TestCategory("UT")]
        public void IsTriggerComboBoxEnableChangedNotification()
        {
            bool triggerComboBoxEnableChangeNotified = false;
            item.PropertyChanged += (sender, args) =>
            {
                if (args.PropertyName == "IsTriggerTypeComboBoxEnabled")
                {
                    triggerComboBoxEnableChangeNotified = true;
                }
            };
            item.IsTriggerTypeComboBoxEnabled = true;
            Assert.IsTrue(triggerComboBoxEnableChangeNotified);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void IsDummyItemIndicator()
        {
            item.IsCallout = false;
            item.IsCaption = false;
            item.IsVoice = false;
            Assert.IsTrue(item.IsDummyItem);

            item.IsCallout = true;
            Assert.IsFalse(item.IsDummyItem);

            item.IsVoice = true;
            Assert.IsFalse(item.IsDummyItem);
        }

        [TestMethod]
        [TestCategory("UT")]
        public void IsEmptyIndicator()
        {
            item.CaptionText = string.Empty;
            item.CalloutText = string.Empty;
            Assert.IsTrue(item.IsEmpty);

            item.HasShortVersion = true;
            item.CalloutText = "callout text";
            item.CaptionText = string.Empty;
            Assert.IsFalse(item.IsEmpty);

            item.CaptionText = string.Empty;
            item.CalloutText = "callout text";
            item.HasShortVersion = false;
            Assert.IsTrue(item.IsEmpty);
        }

    }
}
