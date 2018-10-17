using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;

namespace RWord
{
    class openWord
    {
        public void oWord(string sourse)
        {
            Word.Document doc = null;

            try
            {
                Word.Application app = new Word.Application();
                doc = app.Documents.Open(sourse, ReadOnly:true);
                doc.Activate();
                Word.Bookmarks wBookmarks = doc.Bookmarks;
                System.Diagnostics.Process.Start(sourse);
            }

            catch (Exception ex)
            {
                MessageBox.Show("Error");
            }
        }


    }
}
