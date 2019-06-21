using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Linq;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Imaging;
using System.Windows.Threading;

using Microsoft.Office.Interop.PowerPoint;

using PowerPointLabs.ActionFramework.Common.Extension;
using PowerPointLabs.ELearningLab.Converters;
using PowerPointLabs.ELearningLab.ELearningWorkspace.Model;
using PowerPointLabs.ELearningLab.ELearningWorkspace.ModelFactory;
using PowerPointLabs.ELearningLab.Extensions;
using PowerPointLabs.ELearningLab.Service;
using PowerPointLabs.ELearningLab.Utility;
using PowerPointLabs.Models;
using PowerPointLabs.TextCollection;
using PowerPointLabs.Utils;
using PowerPointLabs.Utils.Windows;
using PowerPointLabs.Views;

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
                    return Items[0] is ExplanationItem;
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
            InitializeComponent();
            syncImage.Source = CommonUtil.CreateBitmapSource(Properties.Resources.SyncExplanationIcon);
            createImage.Source = CommonUtil.CreateBitmapSource(Properties.Resources.AddExplanationIcon);
            isSynced = true;          
            InitializeBackgroundWorker();
            slide = this.GetCurrentSlide();
            if (slide == null)
            {
                SetupSpinnerOnLoadingBegin(content: ELearningLabText.NoSlideSelectedMessage);
                return;
            }
            SetupSpinnerOnLoadingBegin(content: ELearningLabText.OnLoadingMessage);
            slideId = slide.ID;
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
            slide = this.GetCurrentSlide();
            if (slide == null)
            {
                slideId = -1;
                SetupSpinnerOnLoadingBegin(ELearningLabText.NoSlideSelectedMessage);
                return;
            }
            // We do not re-initailize elearning lab if 
            // the current slide is the same as previous slide. 
            // This can happen when user opens presentation mode on current slide
            // and exit presentation mode subsequently.
            // check if the current slide is the same as previous slide
            if (slide != null && slide.ID.Equals(slideId))
            {
                return;
            }
            // update current slide instance
            slideId = slide.ID;
            isSynced = true;
            listView.ItemsSource = null;
            SetupSpinnerOnLoadingBegin(ELearningLabText.OnLoadingMessage);
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
            if (slide == null)
            {
                isSynced = true;
                return;
            }
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
                DialogResult result = MessageBoxUtil.Show(
                       ELearningLabText.PromptToSyncMessage,
                       ELearningLabText.ELearningTaskPaneLabel, MessageBoxButtons.YesNo);
                if (result == DialogResult.Yes)
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
                    if (item is ExplanationItem)
                    {
                        ExplanationItem selfExplanationClickItem = item as ExplanationItem;
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
            Items = LoadItems(e);
            UpdateClickNumAndTriggerInItems(useWorker: true, e: e);
            foreach (ClickItem item in Items)
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                item.PropertyChanged += ListViewItemPropertyChanged;
            }
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
                    RemoveSpinnerOnLoadingFinished();
                });
            }
        }

        private void SyncClickItems()
        {
            bool removeAzureAudioIfAccountInvalid = false;
            bool removeWatsonAudioIfAccountInvalid = false;
            if (IsAzureVoiceSelected())
            {
                removeAzureAudioIfAccountInvalid = !AzureRuntimeService.IsAzureAccountPresentAndValid;
            }
            if (IsWatsonVoiceSelected())
            {
                removeWatsonAudioIfAccountInvalid = !WatsonRuntimeService.IsWatsonAccountPresentAndValid;
            }
            // also align eLL first item click No to be the same as first click No on animation pane
            SyncCustomAnimationToTaskpane(uncheckAzureAudio: removeAzureAudioIfAccountInvalid,
                uncheckWatsonAudio: removeWatsonAudioIfAccountInvalid);
            List<ExplanationItem> items = Items.Where(x => x is ExplanationItem)
                .Cast<ExplanationItem>().ToList();
            ELearningService eLearningService = new ELearningService(slide, items);
            RemoveLabAnimationsFromAnimationPane(eLearningService);
            // align first click No on animation pane to be the same as eLL first click No
            AlignFirstClickNumberForCustomItem();
            ELearningLabTextStorageService.StoreSelfExplanationTextToSlide(
                Items.Where(x => x is ExplanationItem && !((ExplanationItem)x).IsEmpty)
                .Cast<ExplanationItem>().ToList(), slide);
            SyncLabItemToAnimationPane(eLearningService);
        }
        private ObservableCollection<ClickItem> LoadItems(DoWorkEventArgs e)
        {
            SelfExplanationTagService.Clear();
            int clickNo = FirstClickNumber;
            ObservableCollection<ClickItem> customItems = new ObservableCollection<ClickItem>();
            ObservableCollection<ClickItem> clickItems = new ObservableCollection<ClickItem>();
            List<Dictionary<string, string>> explanationItems =
                ELearningLabTextStorageService.LoadSelfExplanationsFromSlide(slide);
            ClickItem customClickBlock;
            SelfExplanationTagService.PopulateTagNos(slide.GetShapesWithNameRegex(ELearningLabText.PPTLShapeNameRegex)
                .Select(x => x.Name).ToList());
            HashSet<int> tagNums = new HashSet<int>();
            List<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>().ToList();
            int startIdx = 0;
            bool hasReachedEndOfSequence = effects.Count == 0;
            List<CustomEffect> customEffects = new List<CustomEffect>();
            List<ELLEffect> pptlEffects = new List<ELLEffect>();
            do
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return clickItems;
                }
                customEffects = new List<CustomEffect>();
                for (int i = startIdx; i < effects.Count; i++)
                {
                    Effect effect = effects.ElementAt(i);
                    if (i > startIdx && effect.Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick)
                    {
                        startIdx = i;
                        break;
                    }
                    if (i == effects.Count - 1)
                    {
                        hasReachedEndOfSequence = true;
                    }
                    bool isPPTLEffect = SelfExplanationTagService.ExtractTagNo(effect.Shape.Name) != -1;
                    if (!isPPTLEffect)
                    {
                        customEffects.Add(new CustomEffect(effect.Shape.Name, effect.Shape.Id.ToString(),
                            EffectToAnimationTypeConverter.GetAnimationTypeOfEffect(effect)));
                    }
                }
                customClickBlock =
                    new CustomItemFactory(customEffects).GetBlock();
                if (customClickBlock != null)
                {
                    customClickBlock.ClickNo = clickNo;
                    customItems.Add(customClickBlock);
                }               
                clickNo++;
            }
            while (startIdx <= effects.Count - 1 && !hasReachedEndOfSequence);

            while (explanationItems != null && explanationItems.Count > 0 && customItems.Count > 0)
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return clickItems;
                }
                ExplanationItem expItem = CreateExpItemFromDictionary(explanationItems.First());
                CustomItem customItem = customItems.ElementAt(0) as CustomItem;
                if (customItem.ClickNo <= expItem.ClickNo)
                {
                    clickItems.Add(customItem);
                    customItems.RemoveAt(0);
                }
                else
                {
                    clickItems.Add(expItem);
                    explanationItems.RemoveAt(0);
                }
            }
            while (explanationItems != null && explanationItems.Count > 0)
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return clickItems;
                }
                ExplanationItem expItem = CreateExpItemFromDictionary(explanationItems.First());
                explanationItems.RemoveAt(0);
                clickItems.Add(expItem);
            }
            while (customItems.Count > 0)
            {
                if (worker.CancellationPending)
                {
                    e.Cancel = true;
                    return clickItems;
                }
                CustomItem customItem = customItems.ElementAt(0) as CustomItem;
                customItems.RemoveAt(0);
                clickItems.Add(customItem);
            }
            return clickItems;
        }

        #region Custom Event Handlers for SelfExplanationBlockView

        private void HandleUpButtonClickedEvent(object sender, RoutedEventArgs e)
        {
            ExplanationItem labItem = ((Button)e.OriginalSource).CommandParameter as ExplanationItem;
            int index = Items.ToList().FindIndex(x => x is ExplanationItem
            && ((ExplanationItem)x).TagNo == labItem.TagNo);
            if (index > 0)
            {
                Items.Move(index, index - 1);
            }
            UpdateClickNumAndTriggerInItems(useWorker: false, e: null);
            ScrollItemToView(labItem);
            isSynced = false;
        }
        private void HandleDownButtonClickedEvent(object sender, RoutedEventArgs e)
        {
            ExplanationItem labItem = ((Button)e.OriginalSource).CommandParameter as ExplanationItem;
            int index = Items.ToList().FindIndex(x => x is ExplanationItem
            && ((ExplanationItem)x).TagNo == labItem.TagNo);
            if (index < Items.Count() - 1 && index >= 0)
            {
                Items.Move(index, index + 1);
            }
            UpdateClickNumAndTriggerInItems(useWorker: false, e: null);
            ScrollItemToView(labItem);
            isSynced = false;
        }
        private void HandleDeleteButtonClickedEvent(object sender, RoutedEventArgs e)
        {
            ExplanationItem labItem = ((Button)e.OriginalSource).CommandParameter as ExplanationItem;
            int index = Items.ToList().FindIndex(x => x is ExplanationItem
            && ((ExplanationItem)x).TagNo == labItem.TagNo);
            Items.RemoveAt(index);
            UpdateClickNumAndTriggerInItems(useWorker: false, e: null);
            isSynced = false;
        }
        private void HandleTriggerTypeComboBoxSelectionChangedEvent(object sender, RoutedEventArgs e)
        {
            UpdateClickNumAndTriggerInItems(useWorker: false, e: null);
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
            ExplanationItem selfExplanationClickItem = new ExplanationItem(captionText: string.Empty);
            selfExplanationClickItem.tagNo = SelfExplanationTagService.GenerateUniqueTag();
            Items.Add(selfExplanationClickItem);
            isSynced = false;
            UpdateClickNumAndTriggerInItems(useWorker: false, e: null);
            ScrollListViewToEnd();
        }

        private void AddItemAboveContextMenu_Click(object sender, RoutedEventArgs e)
        {
            ClickItem item = ((MenuItem)sender).CommandParameter as ClickItem;
            ExplanationItem selfExplanationClickItem = new ExplanationItem(captionText: string.Empty);
            selfExplanationClickItem.tagNo = SelfExplanationTagService.GenerateUniqueTag();
            int index;
            if (item is ExplanationItem)
            {
                index = Items.ToList().FindIndex(x => x is ExplanationItem
                && ((ExplanationItem)x).TagNo == ((ExplanationItem)item).TagNo);
            }
            else
            {
                index = Items.IndexOf(item);
            }
            Items.Insert(index, selfExplanationClickItem);
            isSynced = false;
            UpdateClickNumAndTriggerInItems(useWorker: false, e: null);
            ScrollItemToView(selfExplanationClickItem);
        }

        private void AddItemBelowContextMenu_Click(object sender, RoutedEventArgs e)
        {
            ClickItem item = ((MenuItem)sender).CommandParameter as ClickItem;
            ExplanationItem selfExplanationClickItem = new ExplanationItem(captionText: string.Empty);
            selfExplanationClickItem.tagNo = SelfExplanationTagService.GenerateUniqueTag();
            int index;
            if (item is ExplanationItem)
            {
                index = Items.ToList().FindIndex(x => x is ExplanationItem
                && ((ExplanationItem)x).TagNo == ((ExplanationItem)item).TagNo);
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
            UpdateClickNumAndTriggerInItems(useWorker: false, e: null);
            ScrollItemToView(selfExplanationClickItem);
        }

        #endregion

        #region Helper Methods

        private bool IsInSync()
        {
            return isSynced;
        }

        private void SyncCustomAnimationToTaskpane(bool uncheckAzureAudio, bool uncheckWatsonAudio)
        {
            Queue<CustomItem> customClickItems = LoadCustomClickItems();
            ReplaceCustomItemsInItemsSource(customClickItems);
            UpdatePropertiesInItemsSource(uncheckAzureAudio: uncheckAzureAudio, uncheckWatsonAudio: uncheckWatsonAudio);
        }

        private void SyncLabItemToAnimationPane(ELearningService service)
        {
            int totalSelfExplanationItemsCount = service.GetExplanationItemsCount();
            ProcessingStatusForm progressBarForm =
                new ProcessingStatusForm(totalSelfExplanationItemsCount, BackgroundWorkerType.ELearningLabService, service);
            progressBarForm.ShowDialog();
        }

        /// <summary>
        /// This method aligns starting click number between e-learning lab and animation pane.
        /// This is necessary when first click item on e-learning lab pane is self-explanation item.
        /// </summary>
        private void AlignFirstClickNumberForCustomItem()
        {
            // effects contain only custom effects now
            IEnumerable<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>();
            int clickNo = GetClickNoOfFirstCustomItem();
            if (clickNo > 0 && effects.Count() > 0)
            {
                effects.ElementAt(0).Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerOnPageClick;
            }
            else if (clickNo == 0 && effects.Count() > 0)
            {
                effects.ElementAt(0).Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerWithPrevious;
            }
        }

        private int GetClickNoOfFirstCustomItem()
        {
            for (int i = 0; i < Items.Count; i++)
            {
                if (Items[i] is CustomItem)
                {
                    return Items[i].ClickNo;
                }
            }
            return -1;
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

        private void RemoveLabAnimationsFromAnimationPane(ELearningService service)
        {
            service.RemoveLabAnimationsFromAnimationPane();
        }

        private void UpdateClickNumOnClickItem(ClickItem clickItem, int startClickNo, int index)
        {
            bool isOnClickSelfExplanationAfterCustomItem = index > 0 &&
                clickItem is ExplanationItem && (Items.ElementAt(index - 1) is CustomItem)
                && (clickItem as ExplanationItem).TriggerIndex != (int)TriggerType.OnClick;
            bool isFirstOnClickSelfExplanationItem = index == 0
                && (clickItem is ExplanationItem)
                && (clickItem as ExplanationItem).TriggerIndex == (int)TriggerType.OnClick;
            bool isFirstWithPreviousSelfExplanationItem =
                index == 0
                && (clickItem is ExplanationItem)
                && (clickItem as ExplanationItem).TriggerIndex != (int)TriggerType.OnClick;
            bool isDummySelfExplanationItem =
                clickItem is ExplanationItem && (clickItem as ExplanationItem).IsDummyItem;
            /* This commented piece of code is trying to handle the case when first self explanation item (SEI)
             * is dummy item, but the second one is active SEI item.
            bool isAfterDummySelfExplanationItem =
                index > 0 && (Items.ElementAt(index - 1) is ExplanationItem)
                && (Items.ElementAt(index - 1) as ExplanationItem).IsDummyItem;
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

        private void UpdateTriggerOnExplanationItem(ExplanationItem selfExplanationClickItem, int index)
        {
            if ((index > 0 && Items.ElementAt(index - 1) is CustomItem) || index == 0)
            {
                selfExplanationClickItem.IsTriggerTypeComboBoxEnabled = true;
            }
            else
            {
                selfExplanationClickItem.IsTriggerTypeComboBoxEnabled = false;
                selfExplanationClickItem.TriggerIndex = (int)TriggerType.OnClick;
            }
        }

        private void UpdateSelfExplanationItem(ExplanationItem item, bool uncheckAzureAudio, bool uncheckWatsonAudio)
        {
            if (string.IsNullOrEmpty(item.CaptionText.Trim()))
            {
                item.IsVoice = false;
                item.IsCaption = false;
                if (!item.HasShortVersion)
                {
                    item.IsCallout = false;
                }
            }
            if (item.HasShortVersion && string.IsNullOrEmpty(item.CalloutText.Trim()))
            {
                item.IsCallout = false;
                item.HasShortVersion = false;
            }
            if (item.CaptionText.Trim().Equals(item.CalloutText.Trim()))
            {
                item.HasShortVersion = false;
            }
            if ((uncheckAzureAudio && AudioService.IsAzureVoiceSelectedForItem(item))
                || (uncheckWatsonAudio && AudioService.IsWatsonVoiceSelectedForItem(item)))
            {
                item.IsVoice = false;
                item.VoiceLabel = string.Empty;
            }
        }

        /// <summary>
        /// Load custom animations from animation pane separated by click
        /// </summary>
        /// <returns>Queue of CustomClickItem</returns>
        private Queue<CustomItem> LoadCustomClickItems()
        {
            int clickNo = FirstClickNumber;
            Queue<CustomItem> customItems = new Queue<CustomItem>();
            List<Effect> effects = slide.TimeLine.MainSequence.Cast<Effect>().ToList();
            int startIdx = 0;
            bool hasReachedEndOfSequence = effects.Count == 0;
            do
            {
                List<CustomEffect> customEffects = new List<CustomEffect>();
                for (int i = startIdx; i < effects.Count; i++)
                {
                    Effect effect = effects.ElementAt(i);
                    if (i > startIdx && effect.Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick)
                    {
                        startIdx = i;
                        break;
                    }
                    if (i == effects.Count - 1)
                    {
                        hasReachedEndOfSequence = true;
                    }
                    bool isCustomEffect = SelfExplanationTagService.ExtractTagNo(effect.Shape.Name) == -1;
                    if (isCustomEffect)
                    {
                        if (customEffects.Count == 0 && clickNo > 0)
                        {
                            effect.Timing.TriggerType = MsoAnimTriggerType.msoAnimTriggerOnPageClick;
                        }
                        customEffects.Add(new CustomEffect(effect.Shape.Name, effect.Shape.Id.ToString(),
                            EffectToAnimationTypeConverter.GetAnimationTypeOfEffect(effect)));
                    }
                }
                CustomItem customClickBlock =
                    new CustomItemFactory(customEffects).GetBlock() as CustomItem;
                if (customClickBlock != null)
                {
                    customClickBlock.ClickNo = clickNo;
                    customItems.Enqueue(customClickBlock);
                }
                clickNo++;
            }
            while (startIdx <= effects.Count - 1 && !hasReachedEndOfSequence);
            return customItems;
        }

        /// <summary>
        /// Replace all CustomClickItem in ItemsSource with `customClickItems`
        /// If there are more CustomClickItem in ItemsSource, those are deleted.
        /// Additional customClickItem left in customClickItems will be appended to the back of list.
        /// </summary>
        /// <param name="customClickItems"></param>
        /// <returns></returns>
        private ObservableCollection<ClickItem> ReplaceCustomItemsInItemsSource(Queue<CustomItem> customClickItems)
        {
            for (int i = 0; i < Items.Count(); i++)
            {
                ClickItem clickItem = Items.ElementAt(i);
                if (clickItem is CustomItem)
                {
                    if (customClickItems.Count() > 0)
                    {
                        CustomItem customClickItem = customClickItems.Dequeue();
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
                CustomItem customClickItem = customClickItems.Dequeue();
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
        private ObservableCollection<ClickItem> UpdatePropertiesInItemsSource(bool uncheckAzureAudio, bool uncheckWatsonAudio)
        {
            int clickNo = FirstClickNumber;
            for (int i = 0; i < Items.Count(); i++)
            {
                ClickItem clickItem = Items.ElementAt(i);
                UpdateClickNumOnClickItem(clickItem, clickNo, i);
                if (clickItem is ExplanationItem)
                {
                    UpdateSelfExplanationItem(clickItem as ExplanationItem, uncheckAzureAudio, uncheckWatsonAudio);
                }
            }
            return Items;
        }

        private void UpdateClickNumAndTriggerInItems(bool useWorker, DoWorkEventArgs e)
        {
            int clickNum = FirstClickNumber;
            for (int i = 0; i < Items.Count(); i++)
            {
                if (useWorker && worker.CancellationPending)
                {
                    e.Cancel = true;
                    return;
                }
                ClickItem clickItem = Items.ElementAt(i);
                UpdateClickNumOnClickItem(clickItem, clickNum, i);
                if (clickItem is ExplanationItem)
                {
                    UpdateTriggerOnExplanationItem(clickItem as ExplanationItem, i);
                }
            }
        }

        private bool IsAzureVoiceSelected()
        {
            foreach (ClickItem item in Items)
            {
                if (item is ExplanationItem && AudioService.IsAzureVoiceSelectedForItem(item as ExplanationItem))
                {
                    return true;
                }
            }
            return false;
        }

        private bool IsWatsonVoiceSelected()
        {
            foreach (ClickItem item in Items)
            {
                if (item is ExplanationItem && AudioService.IsWatsonVoiceSelectedForItem(item as ExplanationItem))
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

        private void SetupSpinnerOnLoadingBegin(string content)
        {
            createButton.IsEnabled = false;
            syncButton.IsEnabled = false;
            eLLPane.Visibility = Visibility.Collapsed;
            loadingPane.Visibility = Visibility.Visible;
            loadingLabel.Content = content;
            createImage.Opacity = 0.5;
            syncImage.Opacity = 0.5;
        }

        private void RemoveSpinnerOnLoadingFinished()
        {
            createButton.IsEnabled = true;
            syncButton.IsEnabled = true;
            loadingPane.Visibility = Visibility.Collapsed;
            eLLPane.Visibility = Visibility.Visible;
            createImage.Opacity = 1;
            syncImage.Opacity = 1;
        }

        private ExplanationItem CreateExpItemFromDictionary(Dictionary<string, string> expItemDic)
        {
            ExplanationItem expItem = new ExplanationItem(expItemDic[ELearningLabText.CaptionTextIdentifier],
                                    expItemDic[ELearningLabText.CalloutTextIdentifier], expItemDic[ELearningLabText.VoiceLabel],
                                    expItemDic[ELearningLabText.CalloutIdentifier] == "Y", expItemDic[ELearningLabText.CaptionIdentifier] == "Y",
                                    expItemDic[ELearningLabText.AudioIdentifier] == "Y", tagNo: Convert.ToInt32(expItemDic[ELearningLabText.TagNoIdentifier]));
            expItem.ClickNo = Convert.ToInt32(expItemDic[ELearningLabText.ClickNumIdentifier]);
            expItem.TriggerIndex = expItemDic[ELearningLabText.TriggerOnClick] == "Y" ? (int)TriggerType.OnClick : (int)TriggerType.WithPrevious;
            if (expItem.IsDummyItem)
            {
                expItem.tagNo = SelfExplanationTagService.GenerateUniqueTag();
            }
            return expItem;
        }

        #endregion
    }
}
