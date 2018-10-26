using System;
using System.Collections.Generic;
using System.IO;
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
                //doc.Close();
            }
            
            catch (Exception ex)
            {
                MessageBox.Show("Error");
            }
        }
       
        private Word.Application wordapp;
        private Word.Document worddocument;
        private string nameBank;
        private string bankBook;
        private string bik;
        private string inn;
        private string kpp;
        private string nameAdresat;
        private string bankAdresat;

        public void OpWord(string file, Form1 form)
        {

            try
            {
                object TM = Type.Missing;
                Object filename = file;
                Object confirmConversions = true;                   //При true в случае открытия документа не формата Word будет выводится диалоговое окно конвертирования файла
                Object readOnly = true;                            //При true документ открывается только для чтения            
                Object addToRecentFiles = false;                     //При true имя открываемого файла добавляется в список недавно открытых файлов в меню Файл.
                Object passwordDocument = TM;             //Пароль открываемого документа если он есть
                Object passwordTemplate = TM;             //Пароль шаблона документа если он есть
                Object revert = false;                              //При true возможно повторное открытие экземпляра того же документа с потерей изменений в открытом ранее. При false новый экземпляр не открывается.
                Object writePasswordDocument = TM;        //Пароль для сохранения документа   
                Object writePasswordTemplate = TM;        //Пароль для сохранения шаблона 
                Object format = TM;                       //Одна из следующих Word.WdOpenFormat констант wdOpenFormatAllWord, wdOpenFormatAuto, wdOpenFormatDocument,  wdOpenFormatEncodedText, wdOpenFormatRTF, wdOpenFormatTemplate, wdOpenFormatText, wdOpenFormatUnicodeText или wdOpenFormatWebPages. По умолчанию wdOpenFormatAuto.
                Object encoding = TM;                     //Кодовая страница, или набор символов, (кодировка) для просмотра документа, Значение по умолчанию - системная кодовая страница. Задается как Microsoft.Office.Core.MsoEncoding.msoEncodingUSASCII;
                Object oVisible = true;                             //При true документ открывается как видимый.
                Object openConflictDocument = TM;
                Object openAndRepair = TM;                //При true делается попытка восстановить поврежденный документ.
                Object documentDirection = TM;            //Направление текста - одна из Word.WdDocumentDirection констант: WdLeftToRight, WdRightToLeft.
                Object noEncodingDialog = false;                    //При true подавляется показ диалогового окна Encoding, которое отображается если кодировка не распознана.
                Object xmlTransform = TM;                 //Определяет тип XML данных при XML преобразованиях 
                wordapp = new Word.Application();                     //Открываем новое приложение Word
                wordapp.Visible = false;                             //Делаем его невидимым
                worddocument = wordapp.Documents.Open(ref filename, ref confirmConversions, ref readOnly, ref addToRecentFiles, ref passwordDocument, ref passwordTemplate, ref revert, ref writePasswordDocument, ref writePasswordTemplate, ref format, ref encoding, ref oVisible, ref openConflictDocument, ref documentDirection, ref noEncodingDialog, ref xmlTransform);    //Открываем нужный документ
                nameBank = worddocument.Range(worddocument.Tables[1].Cell(1, 1).Range.Start, worddocument.Tables[1].Cell(1, 1).Range.End - 1).Text;
                bankBook = worddocument.Range(worddocument.Tables[1].Cell(2, 3).Range.Start, worddocument.Tables[1].Cell(2, 3).Range.End - 1).Text;
                bik = worddocument.Range(worddocument.Tables[1].Cell(1, 3).Range.Start, worddocument.Tables[1].Cell(1, 3).Range.End - 1).Text;
                inn = worddocument.Range(worddocument.Tables[1].Cell(4, 1).Range.Start, worddocument.Tables[1].Cell(4, 1).Range.End - 1).Text;
                kpp = worddocument.Range(worddocument.Tables[1].Cell(4, 3).Range.Start, worddocument.Tables[1].Cell(4, 3).Range.End - 1).Text;
                nameAdresat = worddocument.Range(worddocument.Tables[1].Cell(5,1 ).Range.Start, worddocument.Tables[1].Cell(5, 1).Range.End - 1).Text;
                bankAdresat = worddocument.Range(worddocument.Tables[1].Cell(4, 6).Range.Start, worddocument.Tables[1].Cell(4, 6).Range.End - 1).Text;
                // и т.д. (данные взяли, далее делаем с ними, что хотим)


                form.label4.Text = String.Format("Имя банка: {0}", nameBank);
                form.label3.Text = String.Format("Лицевой счёт: {0}", bankBook);
                form.label5.Text = String.Format("БИК банка: {0}", bik);
                form.label6.Text = String.Format(" {0}", inn);
                form.label7.Text = String.Format(" {0}", kpp);
                form.label8.Text = String.Format("Лицевой счёт: {0}" , bankAdresat);
                form.label9.Text = String.Format("Имя получателя: {0}", nameAdresat);

                form.label4.Visible = true;
                form.label3.Visible = true;
                form.label5.Visible = true;
                form.label6.Visible = true;
                form.label7.Visible = true;
                form.label8.Visible = true;
                form.label9.Visible = true;


            }
            catch (Exception e)
            {
                MessageBox.Show(e.Message + "\nВыберите другой файл"
                    + "\nВозможно искать стоит файлы:\n\"Счёт ######.docx\""
                    , "Таблица не найдена", MessageBoxButtons.OK, MessageBoxIcon.Stop);
            }
            worddocument.Close();
            wordapp.Quit(); // Закрываем Ворд

           

        }
       
    }
}
