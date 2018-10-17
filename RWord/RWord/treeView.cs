using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;
using Microsoft.VisualBasic.FileIO;

namespace RWord
{
    
    class treeView
    {
        DriveInfo[] dr = DriveInfo.GetDrives();
        string[] drives = Environment.GetLogicalDrives();

        
        public bool CreateTree(TreeView treeView)
        {

            bool returnValue = false;

            try
            {
                // Рабочий стол
                TreeNode desktop = new TreeNode();
                desktop.Text = "Рабочий стол";
                desktop.Tag = "Desktop";
                desktop.Nodes.Add("");
                treeView.Nodes.Add(desktop);
                // инфа о дисках
                foreach (DriveInfo drv in DriveInfo.GetDrives())
                {

                    TreeNode fChild = new TreeNode();
                    if (drv.DriveType == DriveType.CDRom) // сд ром
                    {
                        fChild.ImageIndex = 1;
                        fChild.SelectedImageIndex = 1;
                    }
                    else if (drv.DriveType == DriveType.Fixed) // хард
                    {
                        fChild.ImageIndex = 0;
                        fChild.SelectedImageIndex = 0;
                    }
                   
                    fChild.Text = drv.Name ;
                    fChild.Nodes.Add("");
                    treeView.Nodes.Add(fChild);
                    returnValue = true;
                }

            }
            catch (Exception ex)
            {
                returnValue = false;
            }
            return returnValue;

        }
     
        public TreeNode EnumerateDirectory(TreeNode parentNode)
        {

            try
            {
                DirectoryInfo rootDir;

                // заполнение рабочего стола
                Char[] arr = { '\\' };
                string[] nameList = parentNode.FullPath.Split(arr);
                string path = "";

                if (nameList.GetValue(0).ToString() == "Рабочий стол")
                {
                    path = SpecialDirectories.Desktop + "\\";

                    for (int i = 1; i < nameList.Length; i++)
                    {
                        path = path + nameList[i] + "\\";
                    }

                    rootDir = new DirectoryInfo(path);
                }
                // грузим каталоги
                else
                {

                    rootDir = new DirectoryInfo(parentNode.FullPath + "\\");
                }

                parentNode.Nodes[0].Remove();
                foreach (DirectoryInfo dir in rootDir.GetDirectories())
                {

                    TreeNode node = new TreeNode();
                    node.Text = dir.Name;
                    node.ImageIndex = 2;
                    node.SelectedImageIndex = 2;
                    node.Nodes.Add("");
                    parentNode.Nodes.Add(node);
                }
                //грузим док файлы
                foreach (FileInfo file in rootDir.GetFiles("*.doc*"))
                {
                    TreeNode node = new TreeNode();

                    node.Text = file.Name;
                    node.ImageIndex = 3;
                    node.SelectedImageIndex = 3;
                    parentNode.Nodes.Add(node);
                    
                }



            }

            catch (Exception ex)
            {
                //TODO : 
            }

            return parentNode;
        }





    }

    
}


