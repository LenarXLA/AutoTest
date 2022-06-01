using System;
using System.Data;
using System.Linq;

using System.Windows.Forms;

namespace AutoTest
{
    public partial class Form1 : Form
    {
        private TestRequest testRequest;
        private string[] strArray;
        public Form1()
        {
            InitializeComponent();

            saveFileDialog1.Filter = ".xls Files (*.xls)|*.xls";
            saveFileDialog1.OverwritePrompt = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            btnAnalyse.Enabled = false;

            //Открываем файл Экселя
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                //Создаём приложение.
                Microsoft.Office.Interop.Excel.Application ObjExcel = new Microsoft.Office.Interop.Excel.Application();
                //Открываем книгу.                                                                                                                                                        
                Microsoft.Office.Interop.Excel.Workbook ObjWorkBook = ObjExcel.Workbooks.Open(openFileDialog1.FileName, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "", true, false, 0, true, false, false);
                //Выбираем таблицу(лист).
                Microsoft.Office.Interop.Excel.Worksheet ObjWorkSheet;
                ObjWorkSheet = (Microsoft.Office.Interop.Excel.Worksheet)ObjWorkBook.Sheets[1];

                //Очищаем от старого текста окно вывода.
                richTextList.Clear();

                //Выбираем область таблицы
                Microsoft.Office.Interop.Excel.Range usedColumn = ObjWorkSheet.UsedRange.Columns[Int32.Parse(textColumn.Text)];
                usedColumn = usedColumn.Rows[textBoxBefore.Text +":"+ textBoxAfter.Text];
                //Добавляем в массив
                Array myvalues = (Array)usedColumn.Cells.Value2;
                strArray = myvalues.OfType<object>().Select(o => o.ToString()).ToArray();
                //Добавляем полученный из ячейки текст
                foreach(string elem in strArray)
                {
                    richTextList.Text = richTextList.Text + elem + "\n";
                }


                //это чтобы форма прорисовывалась (не подвисала)...
                Application.DoEvents();
                //Удаляем приложение (выходим из экселя)
                ObjExcel.Quit();
            }
            btnAnalyse.Enabled = true;
        }

        private void btnAnalyse_Click(object sender, EventArgs e)
        {
            testRequest = new TestRequest();
            testRequest.Setup();

            foreach (string el in strArray)
            {
                testRequest.FindElementError(el);
            }

            testRequest.Quit();

            //Очищаем от старого текста окно вывода.
            richTextListResults.Clear();

            //Добавляем полученный текст
            foreach (string elem in testRequest.results)
            {
                richTextListResults.Text = richTextListResults.Text + elem + "\n";
            }
        }

        private void btnSaveToExcel_Click(object sender, EventArgs e)
        {
            if (saveFileDialog1.ShowDialog() == DialogResult.Cancel)
                return;
            // получаем выбранный файл
            string filename = saveFileDialog1.FileName;
            // сохраняем текст в файл
            System.IO.File.WriteAllText(filename, richTextListResults.Text);
            MessageBox.Show("Файл сохранен");

        }
    }
}
