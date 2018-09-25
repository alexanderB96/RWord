using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.ComponentModel;
using System.IO;
using System.Windows.Forms;

namespace RWord
{
    
    class treeView
    {
        DriveInfo[] dr = DriveInfo.GetDrives();
        public void Open(Form1 form)
        {
          
            foreach (DriveInfo d in dr)
            {
                try
                {
                    DirectoryInfo dir = new DirectoryInfo(d.Name);
                    TreeNode aNode = new TreeNode(d.Name);
                    foreach (DirectoryInfo drs in dir.GetDirectories())
                    {
                        aNode.Nodes.Add(drs.Name);
                    }
                    form.Provodnik.Nodes.Add(aNode);
                }
                catch (Exception ex)
                {
                    form.label1.Text = String.Format("Ошибка инициализации диска: {0}", d);
                }
            }
        }

    }

    
}


