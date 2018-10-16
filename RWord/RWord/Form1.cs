using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using System.Reflection;
using System.IO;

namespace RWord
{
    public partial class Form1 : Form
    {
        treeView fe = new treeView();
        public Form1()
        {
            InitializeComponent();
            
            fe.CreateTree(this.Provodnik);

        }

        private void filepyt_MouseClick(object sender, MouseEventArgs e)
        {
            OpenFileDialog File = new OpenFileDialog();
            File.Title = "Выбереите";
            File.Filter = "doc files 2003| *.doc| docx files 2007 |*.docx";
            if (File.ShowDialog() == DialogResult.OK)
            {
                filepyt.Text = File.FileName;
                filepyt.TextAlign = HorizontalAlignment.Left;
                filepyt.Enabled = false;
            }

            openWord oW = new openWord();
            oW.oWord(filepyt.Text);
        }

        private void Provodnik_AfterSelect(object sender, TreeViewEventArgs e)
        {
          var node = Provodnik.SelectedNode;
            try
            {
                var pt = node.FullPath;
                label1.Text = String.Format("{0}", pt);
            }

            catch
            {
                label1.Text = " Путь не определён ";
            }
            

        }

        private void Provodnik_BeforeExpand(object sender, TreeViewCancelEventArgs e)
        {
            if (e.Node.Nodes[0].Text == "")
            {
                TreeNode node = fe.EnumerateDirectory(e.Node);
            }
           
        }

        private void button1_Click(object sender, EventArgs e) //открытие выбранного файла
        {
            openWord oW = new openWord();
            oW.oWord(label1.Text);
        }

        private void Provodnik_Click(object sender, EventArgs e)
        {
            // не робит что-то
            /*if (label1.Text != "*.doc*")
                 button1.Enabled = false;
            else
                button1.Enabled = true;*/
        }
    }
}
