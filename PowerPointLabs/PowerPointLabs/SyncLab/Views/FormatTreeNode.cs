using System;

using PowerPointLabs.SyncLab.ObjectFormats;

namespace PowerPointLabs.SyncLab.Views
{
    public class FormatTreeNode
    {
        private readonly string name;
        private FormatTreeNode parentNode = null;
        private bool? isChecked = false;
        // only one of childrenNodes and format is a null value
        private FormatTreeNode[] childrenNodes = null;
        private Format format = null;

        public FormatTreeNode(String name, Format format)
        {
            this.name = name;
            this.format = format;
        }

        public FormatTreeNode(String name, params FormatTreeNode[] childrenNodes)
        {
            this.name = name;
            for (int i = 0; i < childrenNodes.Length; i++)
            {
                childrenNodes[i].parentNode = this;
            }
            this.childrenNodes = childrenNodes;
        }

        public string Name
        {
            get
            {
                return name;
            }
        }

        public bool? IsChecked
        {
            get
            {
                return isChecked;
            }
            set
            {
                isChecked = value;
            }
        }

        public FormatTreeNode ParentNode
        {
            get
            {
                return parentNode;
            }
            set
            {
                parentNode = value;
            }
        }

        public FormatTreeNode[] ChildrenNodes
        {
            get
            {
                return childrenNodes;
            }
        }

        public Format Format
        {
            get
            {
                return format;
            }
        }

        public bool IsCategoryNode
        {
            get
            {
                return childrenNodes != null;
            }
        }

        public bool IsFormatNode
        {
            get
            {
                return format != null;
            }
        }

        public FormatTreeNode Clone()
        {
            return Clone(null);
        }

        public FormatTreeNode Clone(FormatTreeNode parent)
        {
            FormatTreeNode cloned = null;
            if (this.IsFormatNode)
            {
                cloned = new FormatTreeNode(name, format);
            }
            else
            {
                FormatTreeNode[] clonedChildren = new FormatTreeNode[childrenNodes.Length];
                for (int i = 0; i < clonedChildren.Length; i++)
                {
                    clonedChildren[i] = childrenNodes[i].Clone();
                }
                cloned = new FormatTreeNode(name, clonedChildren);
            }
            return cloned;
        }

    }
}
