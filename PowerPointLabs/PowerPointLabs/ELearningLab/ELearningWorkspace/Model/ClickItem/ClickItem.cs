﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;

using PowerPointLabs.ActionFramework.Common.Log;

namespace PowerPointLabs.ELearningLab.ELearningWorkspace.Model
{
    public class ClickItem : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;

        public int ClickNo
        {
            get
            {
                return clickNo;
            }
            set
            {
                clickNo = value;
                NotifyPropertyChanged("ClickNo");
            }
        }

        public bool ShouldLabelDisplay
        {
            get
            {
                if (this is CustomClickItem)
                {
                    return true;
                }
                else
                {
                    SelfExplanationClickItem selfExplanationClickItem = this as SelfExplanationClickItem;
                    bool isItemVolunteerOnClick = !selfExplanationClickItem.IsDummyItem
                        && selfExplanationClickItem.IsTriggerTypeComboBoxEnabled
                        && selfExplanationClickItem.TriggerIndex == (int)TriggerType.OnClick;
                    bool isItemForcedOnClick = !selfExplanationClickItem.IsDummyItem
                        && !selfExplanationClickItem.IsTriggerTypeComboBoxEnabled;
                    bool isFirstItemAndStartWithPrevious = !selfExplanationClickItem.IsDummyItem
                        && selfExplanationClickItem.IsTriggerTypeComboBoxEnabled
                        && selfExplanationClickItem.TriggerIndex == (int)TriggerType.WithPrevious
                        && selfExplanationClickItem.ClickNo == 0;                  
                    return isItemVolunteerOnClick || isItemForcedOnClick || isFirstItemAndStartWithPrevious;
                }
            }
        }
        public ClickItem()
        { }
        private int clickNo;

        public override bool Equals(object other)
        {
            if (other == null || other.GetType() != GetType())
            {
                return false;
            }

            if (ReferenceEquals(other, this))
            {
                return true;
            }

            return Equals(this, other);
        }

        public override int GetHashCode()
        {
            var hashCode = 2147116840;
            hashCode = hashCode * -1521134295 + ClickNo.GetHashCode();
            hashCode = hashCode * -1521134295 + clickNo.GetHashCode();
            return hashCode;
        }

        public void NotifyPropertyChanged([CallerMemberName] string propertyName = "")
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

    }
}
