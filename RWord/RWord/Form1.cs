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
        public Form1()
        {
            InitializeComponent();

            treeView tree = new treeView();
            tree.Open(this);
          
        }

        private void filepyt_MouseClick(object sender, MouseEventArgs e)
        {
            OpenFileDialog File = new OpenFileDialog();
            File.Title = "Выбереите";
            File.Filter = "doc files 2003| *.doc| docx files 2007 |*.docx";
            if(File.ShowDialog() == DialogResult.OK)
            {
                filepyt.Text = File.FileName;
                filepyt.TextAlign = HorizontalAlignment.Left;
                filepyt.Enabled = false;
            }

            openWord oW = new openWord();
            oW.oWord(filepyt.Text);
        }
    }
}
