using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Controls;

namespace Excel2Oracle
{
    public enum NodeType
    {
        File = 0,
        Directory,
        Root
    }
    /// <summary>
    /// TreeViewHelper
    /// </summary>
    public class TemplateTreeViewItem
    {
        public TemplateTreeViewItem(NodeType NodeType,string Path, TemplateTreeViewItem[] Children,string DisplayName)
        {
            this.NodeType = NodeType;
            this.Path = Path;
            this.Children = Children;
            this.DisplayName = DisplayName;
            this.Visiblity = "Collapsed";
        }
        public bool IsExpanded { get; set; }
        public string Icon { get; set; }
        //public string EditIcon { get; set; }
        public string Visiblity { get; set; }
        public int id { get; set; }
        public int parentId { get; set; }
        public NodeType NodeType
        {
            get;
            private set;
        }
        public string Path
        {
            get;
            private set;
        }
        public TemplateTreeViewItem[] Children
        {
            get;
            private set;
        }

        public string DisplayName
        {
            get;
            private set;
        }
    }
}
