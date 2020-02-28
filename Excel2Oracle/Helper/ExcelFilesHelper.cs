using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Text;
namespace Excel2Oracle
{
    /// <summary>
    /// ExcelFilesHelper
    /// </summary>
    public class ExcelFilesHelper
    {
        public static TemplateTreeViewItem[] GetAllExcelFiles(string path)
        {
            List<TemplateTreeViewItem> result = new List<TemplateTreeViewItem>() { };
            string[] allFilesAndDirectories = Directory.GetFileSystemEntries(path);
            foreach (string item in allFilesAndDirectories)
            {
               
                if (IsFile(item))
                {
                    TemplateTreeViewItem treeViewItem = new TemplateTreeViewItem(NodeType.File, item, null,Path.GetFileName(item));
                    treeViewItem.Icon = "Images/excel.jpg";
                    treeViewItem.IsExpanded = false;
                    result.Add(treeViewItem);
                }
                else if (IsDirectory(item))
                {
                    TemplateTreeViewItem[] children = GetAllExcelFiles(item);
                    TemplateTreeViewItem treeViewItem = new TemplateTreeViewItem(NodeType.Directory, item, children, Path.GetFileName(item));
                    treeViewItem.Icon = "Images/dir.jpg";
                    treeViewItem.IsExpanded = true;
                    result.Add(treeViewItem);
                }
            }
            return result.ToArray();
        }

        public static bool IsFile(string path)
        {
            if (File.Exists(path))
            {
                return true;
            }
            return false;
        }
        public static bool IsDirectory(string path)
        {
            if (Directory.Exists(path))
            {
                return true;
            }

            return false;
        }

        
    }
}
