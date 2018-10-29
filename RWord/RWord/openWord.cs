using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Reflection;

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
                doc = app.Documents.Open(sourse, ReadOnly: true);
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
        Excel.Application ObjExcel = new Excel.Application();
        Excel.Worksheet ObjWorkSheet = new Excel.Worksheet();
        Excel.Workbook ObjWorkBook;
        private Excel.Range excelcellsOt;
        private Excel.Range excelcellsDo;
        

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
                nameAdresat = worddocument.Range(worddocument.Tables[1].Cell(5, 1).Range.Start, worddocument.Tables[1].Cell(5, 1).Range.End - 1).Text;
                bankAdresat = worddocument.Range(worddocument.Tables[1].Cell(4, 6).Range.Start, worddocument.Tables[1].Cell(4, 6).Range.End - 1).Text;
                // и т.д. (данные взяли, далее делаем с ними, что хотим)


                form.label4.Text = String.Format("Имя банка: {0}", nameBank);
                form.label3.Text = String.Format("Коррекционный счёт: {0}", bankBook);
                form.label5.Text = String.Format("БИК банка: {0}", bik);
                form.label6.Text = String.Format("{0}", inn);
                form.label7.Text = String.Format("{0}", kpp);
                form.label8.Text = String.Format("Расчётный счёт: {0}", bankAdresat);
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

        public void OpExcel(string file, Form1 form)
        {
            try
            {


                //открытие файла
                ObjWorkBook = ObjExcel.Workbooks.Open(file, Type.Missing, true, Type.Missing, "", "", Type.Missing, Type.Missing, Type.Missing,
                                                                                                       Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                                                                                                                                  Type.Missing, Type.Missing);


                //выбор листа 
                ObjWorkSheet = ObjWorkBook.Sheets[1];

                //поиск по слову Дата
                excelcellsOt = ObjWorkSheet.Cells.Find("Дата", Missing.Value, Missing.Value, Excel.XlLookAt.xlPart, Missing.Value,
                   Excel.XlSearchDirection.xlNext,
                   Missing.Value, Missing.Value, Missing.Value);

                //поиск по слову Итого
                excelcellsDo = ObjWorkSheet.Cells.Find("Итого", Missing.Value, Missing.Value, Excel.XlLookAt.xlPart, Missing.Value,
                   Excel.XlSearchDirection.xlNext,
                   Missing.Value, Missing.Value, Missing.Value);

                //общее пояснение
                //AdrOt - ячейка поиска от
                //AdrDo - ячейка поиска до

                //полученные адреса разделяем
                string[] AdrOtTmp = excelcellsOt.Address.Split('$');
                string[] AdrDoTmp = excelcellsDo.Address.Split('$');

                //изменяем номера строк(чтобы не попадали слова Дата и Итого)
                int AdrOtTmp2 = Convert.ToInt32(Convert.ToInt32(AdrOtTmp[2]) + 1);
                int AdrDoTmp2 = Convert.ToInt32(Convert.ToInt32(AdrDoTmp[2]) - 1);

                //склеиваем обрадно полученные диапозон без Даты и Итого
                string AdrOt = Convert.ToString(AdrOtTmp[1] + AdrOtTmp2);
                string AdrDo = Convert.ToString(AdrDoTmp[1] + AdrDoTmp2);

                //собственно сам диапазон 
                // var numCol = String.Format("{0}:{1}", excelcellsOt.Address, excelcellsDo.Address); //изначально было так, без бубна с переприсвоением
                var numCol = String.Format("{0}:{1}", AdrOt, AdrDo); // с бубном так

                //задаем диапазон поиска
                Excel.Range usedColumn = ObjWorkSheet.Range[numCol];

                Array myvalues = (Array)usedColumn.Cells.Value2;
            // получили массив с датами формата стринг
                string[] strArrayPol = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();
            //преобразовали эти даты в привычный вид
                string[] strArray = strArrayPol.Select(x => DateTime.FromOADate(Convert.ToDouble(x)).ToShortDateString()).ToArray();


            for (int i = 0; i < strArray.Length; i++)
                {
                 form.listBox1.Items.Add(strArray[i]);
                }
           }
            catch (Exception e)
            {
                MessageBox.Show(e.StackTrace, e.Message);

            }
            ObjWorkBook.Close();
            ObjExcel.Quit();
            
        }

    }
}

