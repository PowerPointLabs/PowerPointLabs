using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Collections.Specialized;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Threading;
using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ActionFramework.Common.Log;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.ELearningLab.ELearningWorkspace.ModelFactory;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.ELearningLab.Utility;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;

namespace PowerPointLabs.ELearningLab.Views
{
    /// <summary>
    /// Interaction logic for ELearningLabMainPanel.xaml
    /// </summary>
#pragma warning disable 618
    public partial class ELearningLabMainPanel : UserControl
    {
        public ObservableCollection<ClickItem> Items { get; set; } = new ObservableCollection<ClickItem>();
        public int FirstClickNumber
        {
            get
            {
                return slide.IsFirstAnimationTriggeredByClick() ? 1 : 0;
            }
        }
        public bool IsFirstItemSelfExplanation
        {
            get
            {
                if (Items.Count > 0)
                {
                    return Items[0] is SelfExplanationClickItem;
                }
                return false;
            }
        }

        private PowerPointSlide slide;
        private int slideId;
        private bool isSynced;
        private BackgroundWorker worker;

        public ELearningLabMainPanel()
        {
            slide = this.GetCurrentSlide();
            if (slide == null)
            {
                return;
            }
            slideId = slide.ID;
            InitializeComponent();
            syncImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
               Properties.Resources.SyncExplanationIcon.GetHbitmap(),
               IntPtr.Zero,
               Int32Rect.Empty,
               BitmapSizeOptions.FromEmptyOptions());
            createImage.Source = System.Windows.Interop.Imaging.CreateBitmapSourceFromHBitmap(
               Properties.Resources.AddExplanationIcon.GetHbitmap(),
               IntPtr.Zero,
               Int32Rect.Empty,
               BitmapSizeOptions.FromEmptyOptions());
            isSynced = true;
            InitializeBackgroundWorker();
            worker.RunWorkerAsync();
        }

        public void ListViewItemPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            isSynced = false;
        }

        /// <summary>
        /// This method is called when slide selection is changed 
        /// and e-learning lab is open
        /// </summary>
        public void ReloadELearningLabOnSlideSelectionChanged()
        {
            PowerPointSlide _slide = this.GetCurrentSlide();
            // We do not re-initailize elearning lab if 
            // the current slide is the same as previous slide. 
            // This can happen when user opens presentation mode on current slide
            // and exit presentation mode subsequently.
            if (_slide == null)
            {
                return;
            }
            // check if the current slide is the same as previous slide
            slide = _slide;
            if (_slide.ID.Equals(slideId))
            {
                return;
            }
            // update current slide instance
            slideId = slide.ID;
            isSynced = true;
            listView.ItemsSource = null;
            if (worker.IsBusy)
            {
                worker.CancelAsync();
            }
            else
            {
                worker.RunWorkerAsync();
            }
        }

        /// <summary>
        /// This method is called when slide selection is changed
        /// regardless of whether e-learning lab is open
        /// </summary>
        public void SyncElearningLabOnSlideSelectionChanged()
        {
            // do not check for sync if previous slide is deleted
            try
            {
                int id = slide.ID;
            }
            catch
            {
                isSynced = true;
                return;
            }
            // We do not check for sync if the current slide is the same as previous slide. 
            PowerPointSlide _slide = this.GetCurrentSlide();
            if (_slide != null && _slide.ID.Equals(slideId))
            {
                return;
            }
            if (!IsInSync())
            {
                System.Windows.Forms.DialogResult result = System.Windows.Forms.MessageBox.Show(
                       ELearningLabText.PromptToSyncMessage,
                       ELearningLabText.ELearningTaskPaneLabel, System.Windows.Forms.MessageBoxButtons.YesNo);
                if (result == System.Windows.Forms.DialogResult.Yes)
                {
                    SyncClickItems();
                }
                isSynced = true;
            }
        }

        public void RefreshVoiceLabelOnAudioSettingChanged()
        {
            if (Visibility == Visibility.Visible)
            {
                ObservableCollection<ClickItem> clickItems = listView.ItemsSource as ObservableCollection<ClickItem>;
                foreach (ClickItem item in clickItems)
                {
                    if (item is SelfExplanationClickItem)
                    {
                        SelfExplanationClickItem selfExplanationClickItem = item as SelfExplanationClickItem;
                        if (StringUtility.ExtractDefaultLabelFromVoiceLabel(selfExplanationClickItem.VoiceLabel)
                            .Equals(ELearningLabText.DefaultAudioIdentifier))
                        {
                            selfExplanationClickItem.VoiceLabel = string.Format(ELearningLabText.AudioDefaultLabelFormat,
                                AudioSettingService.selectedVoice.ToString());
                        }
                    }
                }
            }
        }

        private void InitializeBackgroundWorker()
        {
            worker = new BackgroundWorker();
            worker.WorkerSupportsCancellation = true;
            worker.DoWork += Worker_DoWorkToReloadElearningLabItems;
            worker.RunWorkerCompleted += Worker_RunWorkerCompleted;
        }

        private void Worker_DoWorkToReloadElearningLabItems(object sender, DoWorkEventArgs e)
        {
            Logger.Log("loading items");
            Items = LoadItems(e);
            Logger.Log("updating click numbers");
            UpdateClickNoAndTriggerTypeInItems(useWorker: true, e: e);
            Logger.Log("attaching events");
            foreach (ClickItem item in Items)
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                item.PropertyChanged += ListViewItemPropertyChanged;
            }
            Logger.Log("returning");
            return;
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            // update UI
            if (e.Cancelled)
            {
                worker.RunWorkerAsync();
            }
            else
            {
                Dispatcher.Invoke(() =>
                {
                    listView.ItemsSource = null;
                    listView.ItemsSource = Items;
                });
            }
        }

        private void SyncClickItems()
        {
            bool removeAzureAudioIfAccountInvalid = false;
            if (IsAzureVoiceSelected())
            {
                removeAzureAudioIfAccountInvalid = !CheckAzureAccountValidity();
            }
            SyncCustomAnimationToTaskpane(uncheckAzureAudio: removeAzureAudioIfAccountInvalid);
            RemoveLabAnimationsFromAnimationPane();
            AlignFirstClickNumber();
            ELearningLabTextStorageService.StoreSelfExplanationTextToSlide(
                Items.Where(x => x is SelfExplanationClickItem && !((SelfExplanationClickItem)x).IsEmpty)
                .Cast<SelfExplanationClickItem>().ToList(), slide);
            SyncLabItemToAnimationPane();
        }
        private ObservableCollection<ClickItem> LoadItems(DoWorkEventArgs e)
        {
            SelfExplanationTagService.Clear();
            int clickNo = FirstClickNumber;
            ObservableCollection<ClickItem> clickBlocks = new ObservableCollection<ClickItem>();
            List<Dictionary<string, string>> selfExplanationTexts =
                ELearningLabTextStorageService.LoadSelfExplanationsFromSlide(slide);
            ClickItem customClickBlock;
            SelfExplanationClickItem selfExplanationClickBlock;
            Dictionary<string, string> selfExplanationText = (selfExplanationTexts == null || selfExplanationTexts.Count() == 0) ?
                null : selfExplanationTexts.First();
            SelfExplanationTagService.PopulateTagNos(slide.GetShapesWithNameRegex(ELearningLabText.PPTLShapeNameRegex)
                .Select(x => x.Name).ToList());
            HashSet<int> tagNums = new HashSet<int>();
            do
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return clickBlocks;
                }
                customClickBlock =
                    new CustomItemFactory(slide.GetCustomEffectsForClick(clickNo), slide).GetBlock();
                selfExplanationClickBlock =
                    new SelfExplanationItemFactory(slide.GetPPTLEffectsForClick(clickNo), slide).GetBlock() as SelfExplanationClickItem;
                // we ignore self explanation item if the same click has already been added.
                // this can happen if user misplaced already generated self explanation item.
                if (selfExplanationClickBlock != null && tagNums.Contains(selfExplanationClickBlock.tagNo))
                {
                    selfExplanationClickBlock = null;
                }
                else if (selfExplanationClickBlock != null)
                {
                    tagNums.Add(selfExplanationClickBlock.tagNo);
                }
                // load any dummy items from text storage on slide
                while (selfExplanationText != null && selfExplanationClickBlock != null &&
                    Convert.ToInt32(selfExplanationText["TagNo"]) != selfExplanationClickBlock.tagNo)
                {
                    SelfExplanationClickItem dummySelfExplanation =
                        new SelfExplanationClickItem(captionText: selfExplanationText["CaptionText"],
                        calloutText: selfExplanationText["CalloutText"]);
                    dummySelfExplanation.tagNo = SelfExplanationTagService.GenerateUniqueTag();
                    clickBlocks.Add(dummySelfExplanation);
                    selfExplanationTexts.RemoveAt(0);
                    selfExplanationText = selfExplanationTexts.Count() == 0 ? null : selfExplanationTexts.First();
                }

                if (customClickBlock != null)
                {
                    customClickBlock.ClickNo = clickNo;
                    clickBlocks.Add(customClickBlock);
                }
                if (selfExplanationClickBlock != null)
                {
                    selfExplanationClickBlock.ClickNo = clickNo;
                    if (customClickBlock == null && selfExplanationClickBlock is SelfExplanationClickItem) // is independent block
                    {
                        (selfExplanationClickBlock as SelfExplanationClickItem).TriggerIndex = (int)TriggerType.OnClick;
                    }
                    try
                    {
                        selfExplanationClickBlock.CaptionText = selfExplanationText["CaptionText"];
                        selfExplanationClickBlock.CalloutText = selfExplanationText["CalloutText"];
                        selfExplanationClickBlock.HasShortVersion =
                            !selfExplanationClickBlock.CaptionText.Equals(selfExplanationClickBlock.CalloutText);
                        selfExplanationTexts.RemoveAt(0);
                        selfExplanationText = selfExplanationTexts.Count() == 0 ? null : selfExplanationTexts.First();
                    }
                    catch
                    {
                        Logger.Log("AnimationPane contains tagNos that are not present in dictionary");
                    }
                    clickBlocks.Add(selfExplanationClickBlock);
                }
                clickNo++;
            }
            while (customClickBlock != null || selfExplanationClickBlock != null);

            // add remaining dummy explanation items from text storage on slide
            while (selfExplanationText != null)
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return clickBlocks;
                }
                SelfExplanationClickItem dummySelfExplanation =
                    new SelfExplanationClickItem(captionText: selfExplanationText["CaptionText"],
                    calloutText: selfExplanationText["CalloutText"]);
                dummySelfExplanation.tagNo = SelfExplanationTagService.GenerateUniqueTag();
                clickBlocks.Add(dummySelfExplanation);
                selfExplanationTexts.RemoveAt(0);
                selfExplanationText = selfExplanationTexts.Count() == 0 ? null : selfExplanationTexts.First();
            }

            return clickBlocks;
        }

        #region Custom Event Handlers for SelfExplanationBlockView

        private void HandleUpButtonClickedEvent(object sender, RoutedEventArgs e)
        {
            SelfExplanationClickItem labItem = ((Button)e.OriginalSource).CommandParameter as SelfExplanationClickItem;
            int index = Items.ToList().FindIndex(x => x is SelfExplanationClickItem 
            && ((SelfExplanationClickItem)x).TagNo == labItem.TagNo);
            if (index > 0)
            {
                Items.Move(index, index - 1);
            }
            UpdateClickNoAndTriggerTypeInItems(useWorker: false, e: null);
            ScrollItemToView(labItem);
            isSynced = false;
        }
        private void HandleDownButtonClickedEvent(object sender, RoutedEventArgs e)
        {
            SelfExplanationClickItem labItem = ((Button)e.OriginalSource).CommandParameter as SelfExplanationClickItem;
            int index = Items.ToList().FindIndex(x => x is SelfExplanationClickItem 
            && ((SelfExplanationClickItem)x).TagNo == labItem.TagNo);
            if (index < Items.Count() - 1 && index >= 0)
            {
                Items.Move(index, index + 1);
            }
            UpdateClickNoAndTriggerTypeInItems(useWorker: false, e: null);
            ScrollItemToView(labItem);
            isSynced = false;
        }
        private void HandleDeleteButtonClickedEvent(object sender, RoutedEventArgs e)
        {
            SelfExplanationClickItem labItem = ((Button)e.OriginalSource).CommandParameter as SelfExplanationClickItem;
            int index = Items.ToList().FindIndex(x => x is SelfExplanationClickItem 
            && ((SelfExplanationClickItem)x).TagNo == labItem.TagNo);
            Items.RemoveAt(index);
            UpdateClickNoAndTriggerTypeInItems(useWorker: false, e: null);
            isSynced = false;
        }
        private void HandleTriggerTypeComboBoxSelectionChangedEvent(object sender, RoutedEventArgs e)
        {
            UpdateClickNoAndTriggerTypeInItems(useWorker: false, e: null);
        }

        #endregion

        #region XMAL-Binded Event Handler

        private void SyncButton_Click(object sender, RoutedEventArgs e)
        {
            SyncClickItems();
            isSynced = true;
        }

        private void CreateButton_Click(object sender, RoutedEventArgs e)
        {
            SelfExplanationClickItem selfExplanationClickItem = new SelfExplanationClickItem(captionText: string.Empty);
            selfExplanationClickItem.tagNo = SelfExplanationTagService.GenerateUniqueTag();
            Items.Add(selfExplanationClickItem);
            isSynced = false;
            UpdateClickNoAndTriggerTypeInItems(useWorker: false, e: null);
            ScrollListViewToEnd();
        }

        private void AddItemAboveContextMenu_Click(object sender, RoutedEventArgs e)
        {
            ClickItem item = ((MenuItem)sender).CommandParameter as ClickItem;
            SelfExplanationClickItem selfExplanationClickItem = new SelfExplanationClickItem(captionText: string.Empty);
            selfExplanationClickItem.tagNo = SelfExplanationTagService.GenerateUniqueTag();
            int index;
            if (item is SelfExplanationClickItem)
            {
                index = Items.ToList().FindIndex(x => x is SelfExplanationClickItem 
                && ((SelfExplanationClickItem)x).TagNo == ((SelfExplanationClickItem)item).TagNo);
            }
            else
            {
                index = Items.IndexOf(item);
            }
            Items.Insert(index, selfExplanationClickItem);
            isSynced = false;
            UpdateClickNoAndTriggerTypeInItems(useWorker: false, e: null);
            ScrollItemToView(selfExplanationClickItem);
        }

        private void AddItemBelowContextMenu_Click(object sender, RoutedEventArgs e)
        {
            ClickItem item = ((MenuItem)sender).CommandParameter as ClickItem;
            SelfExplanationClickItem selfExplanationClickItem = new SelfExplanationClickItem(captionText: string.Empty);
            selfExplanationClickItem.tagNo = SelfExplanationTagService.GenerateUniqueTag();
            int index;
            if (item is SelfExplanationClickItem)
            {
                index = Items.ToList().FindIndex(x => x is SelfExplanationClickItem 
                && ((SelfExplanationClickItem)x).TagNo == ((SelfExplanationClickItem)item).TagNo);
            }
            else
            {
                index = Items.IndexOf(item);
            }
            if (index < listView.Items.Count - 1)
            {
                Items.Insert(index + 1, selfExplanationClickItem);
            }
            else
            {
                Items.Add(selfExplanationClickItem);
            }
            isSynced = false;
            UpdateClickNoAndTriggerTypeInItems(useWorker: false, e: null);
            ScrollItemToView(selfExplanationClickItem);
        }

        #endregion

        #region Helper Methods

        private bool IsInSync()
        {
            return isSynced;
        }

        private void SyncCustomAnimationToTaskpane(bool uncheckAzureAudio)
        {
            Queue<CustomClickItem> customClickItems = LoadCustomClickItems();
            ReplaceCustomItemsInItemsSource(customClickItems);
            UpdatePropertiesInItemsSource(uncheckAzureAudio: uncheckAzureAudio);
        }

        private void SyncLabItemToAnimationPane()
        {
            ELearningService.SyncLabItemToAnimationPane(slide,
                Items.Where(
                    x => x is SelfExplanationClickItem).Cast<SelfExplanationClickItem>().ToList());
        }

        /// <summary>
        /// This method aligns starting click number between e-learning lab and animation pane.
        /// This is necessary when first click item on e-learning lab pane is self-explanation item.
        /// </summary>
        private void AlignFirstClickNumber()
        {
            IEnumerable<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>();
            if (Items.Count > 0)
            {
                int clickNo = Items[0].ClickNo;
                if (clickNo > 0 && effects.Count() > 0)
                {
                    effects.ElementAt(0).Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                }
                else if (clickNo == 0 && effects.Count() > 0)
                {
                    effects.ElementAt(0).Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerWithPrevious;
                }
            }
        }

        private void ScrollItemToView(ClickItem item)
        {
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Loaded,
                new Action(delegate
                {                    
                    listView.ScrollIntoView(item);
                }));
        }

        private void ScrollListViewToEnd()
        {
            Dispatcher.CurrentDispatcher.BeginInvoke(DispatcherPriority.Loaded, new Action(delegate
            {
                if (listView.Items.Count > 0)
                {
                    listView.ScrollIntoView(listView.Items[listView.Items.Count - 1]);
                }
            }));
        }

        private void RemoveLabAnimationsFromAnimationPane()
        {
            slide.RemoveAnimationsForShapeWithPrefix(ELearningLabText.Identifier);
        }

        private void UpdateClickNoOnClickItem(ClickItem clickItem, int startClickNo, int index)
        {
            bool isOnClickSelfExplanationAfterCustomItem = index > 0 &&
                clickItem is SelfExplanationClickItem && (Items.ElementAt(index - 1) is CustomClickItem)
                && (clickItem as SelfExplanationClickItem).TriggerIndex != (int)TriggerType.OnClick;
            bool isFirstOnClickSelfExplanationItem = index == 0
                && (clickItem is SelfExplanationClickItem)
                && (clickItem as SelfExplanationClickItem).TriggerIndex == (int)TriggerType.OnClick;
            bool isFirstWithPreviousSelfExplanationItem =
                index == 0
                && (clickItem is SelfExplanationClickItem)
                && (clickItem as SelfExplanationClickItem).TriggerIndex != (int)TriggerType.OnClick;
            bool isDummySelfExplanationItem =
                clickItem is SelfExplanationClickItem && (clickItem as SelfExplanationClickItem).IsDummyItem;
            /* This commented piece of code is trying to handle the case when first self explanation item (SEI)
             * is dummy item, but the second one is active SEI item.
            bool isAfterDummySelfExplanationItem =
                index > 0 && (Items.ElementAt(index - 1) is SelfExplanationClickItem)
                && (Items.ElementAt(index - 1) as SelfExplanationClickItem).IsDummyItem;
            */
            if (index == 0)
            {
                clickItem.ClickNo = startClickNo;
                if (isFirstOnClickSelfExplanationItem && !isDummySelfExplanationItem)
                {
                    clickItem.ClickNo = 1;
                    clickItem.ShouldLabelDisplay = true;
                }
                if (isFirstWithPreviousSelfExplanationItem)
                {
                    clickItem.ClickNo = 0;
                    clickItem.ShouldLabelDisplay = true;
                }
                if (isDummySelfExplanationItem)
                {
                    clickItem.ClickNo = 0;
                    clickItem.ShouldLabelDisplay = false;
                }
            }
            else if (isOnClickSelfExplanationAfterCustomItem || isDummySelfExplanationItem)
            {
                clickItem.ClickNo = Items.ElementAt(index - 1).ClickNo;
                clickItem.ShouldLabelDisplay = false;
            }
            else
            {
                clickItem.ClickNo = Items.ElementAt(index - 1).ClickNo + 1;
                clickItem.ShouldLabelDisplay = true;
            }
            clickItem.NotifyPropertyChanged("ShouldLabelDisplay");
        }

        private void UpdateTriggerTypeEnabledOnSelfExplanationItem(SelfExplanationClickItem selfExplanationClickItem, int index)
        {
            if ((index > 0 && Items.ElementAt(index - 1) is CustomClickItem) || index == 0)
            {
                selfExplanationClickItem.IsTriggerTypeComboBoxEnabled = true;
            }
            else
            {
                selfExplanationClickItem.IsTriggerTypeComboBoxEnabled = false;
                selfExplanationClickItem.TriggerIndex = (int)TriggerType.OnClick;
            }
        }

        private void UpdateSelfExplanationItem(SelfExplanationClickItem item, bool uncheckAzureAudio)
        {
            if (string.IsNullOrEmpty(item.CaptionText.Trim()))
            {
                item.IsVoice = false;
                item.IsCaption = false;
            }
            if (string.IsNullOrEmpty(item.CalloutText.Trim()))
            {
                item.IsCallout = false;
                item.HasShortVersion = false;
            }
            if (item.CaptionText.Trim().Equals(item.CalloutText.Trim()))
            {
                item.HasShortVersion = false;
            }
            if (uncheckAzureAudio && AudioService.IsAzureVoiceSelectedForItem(item))
            {
                item.IsVoice = false;
                item.VoiceLabel = string.Empty;
            }
        }

        /// <summary>
        /// Load custom animations from animation pane separated by click
        /// </summary>
        /// <returns>Queue of CustomClickItem</returns>
        private Queue<CustomClickItem> LoadCustomClickItems()
        {
            int clickNo = FirstClickNumber;
            Queue<CustomClickItem> customClickItems = new Queue<CustomClickItem>();
            ClickItem customClickBlock;
            do
            {
                customClickBlock =
                    new CustomItemFactory(slide.GetCustomEffectsForClick(clickNo), slide).GetBlock();

                if (customClickBlock != null)
                {
                    customClickBlock.ClickNo = clickNo;
                    customClickItems.Enqueue((CustomClickItem)customClickBlock);
                }

                clickNo++;
            }
            while (slide.TimeLine.MainSequence.FindFirstAnimationForClick(clickNo) != null);

            return customClickItems;
        }

        /// <summary>
        /// Replace all CustomClickItem in ItemsSource with `customClickItems`
        /// If there are more CustomClickItem in ItemsSource, those are deleted.
        /// Additional customClickItem left in customClickItems will be appended to the back of list.
        /// </summary>
        /// <param name="customClickItems"></param>
        /// <returns></returns>
        private ObservableCollection<ClickItem> ReplaceCustomItemsInItemsSource(Queue<CustomClickItem> customClickItems)
        {
            for (int i = 0; i < Items.Count(); i++)
            {
                ClickItem clickItem = Items.ElementAt(i);
                if (clickItem is CustomClickItem)
                {
                    if (customClickItems.Count() > 0)
                    {
                        CustomClickItem customClickItem = customClickItems.Dequeue();
                        Items.RemoveAt(i);
                        Items.Insert(i, customClickItem);
                    }
                    else
                    {
                        Items.RemoveAt(i);
                        i--;
                    }
                }
            }
            while (customClickItems.Count() > 0)
            {
                CustomClickItem customClickItem = customClickItems.Dequeue();
                Items.Add(customClickItem);
            }
            return Items;
        }

        /// <summary>
        /// Update ClickNo property in ClickItem when old CustomClickItem is replaced with new ones.
        /// Note that we cannot reply on `BlockToIndexConverter` to update ClickNo, 
        /// because ListViewItem which invokes `BlockToIndexConverter`, is loaded lazily.
        /// </summary>
        /// <param name="clickItems"></param>
        /// <returns></returns>
        private ObservableCollection<ClickItem> UpdatePropertiesInItemsSource(bool uncheckAzureAudio)
        {
            int clickNo = FirstClickNumber;
            for (int i = 0; i < Items.Count(); i++)
            {
                ClickItem clickItem = Items.ElementAt(i);
                UpdateClickNoOnClickItem(clickItem, clickNo, i);
                if (clickItem is SelfExplanationClickItem)
                {
                    UpdateSelfExplanationItem(clickItem as SelfExplanationClickItem, uncheckAzureAudio);
                }
            }
            return Items;
        }

        private void UpdateClickNoAndTriggerTypeInItems(bool useWorker, DoWorkEventArgs e)
        {
            int clickNo = FirstClickNumber;
            for (int i = 0; i < Items.Count(); i++)
            {
                if (useWorker && worker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                ClickItem clickItem = Items.ElementAt(i);
                UpdateClickNoOnClickItem(clickItem, clickNo, i);
                if (clickItem is SelfExplanationClickItem)
                {
                    UpdateTriggerTypeEnabledOnSelfExplanationItem(clickItem as SelfExplanationClickItem, i);
                }
            }
        }

        private bool CheckAzureAccountValidity()
        {
            AzureAccountStorageService.LoadUserAccount();
            if (!AzureRuntimeService.IsAzureAccountPresent() || !AzureRuntimeService.IsValidUserAccount())
            {
                MessageBox.Show("Azure Account Authentication Failed. \nAzure Voices Cannot Be Generated.");
                return false;
            }
            return true;
        }

        private bool IsAzureVoiceSelected()
        {
            foreach (ClickItem item in Items)
            {
                if (item is SelfExplanationClickItem && AudioService.IsAzureVoiceSelectedForItem(item as SelfExplanationClickItem))
                {
                    return true;
                }
            }
            return false;
        }

        /// <summary>
        /// This method iterates through all slides and check if a slide with slideId exists
        /// We can use slideId to check if a slide exists in slide deck because slideId is 
        /// uniquely assigned to slide upon initialization and does not change with slide.
        /// </summary>
        /// <param name="slideId"></param>
        /// <returns></returns>
        private bool DoesSlideExist(int slideId)
        {
            Slides slides = Globals.ThisAddIn.Application.ActivePresentation.Slides;
            foreach (Slide slide in slides)
            {
                if (slideId == slide.SlideID)
                {
                    return true;
                }
            }
            return false;
        }

        #endregion

    }
}
