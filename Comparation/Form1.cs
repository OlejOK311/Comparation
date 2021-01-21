using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;

namespace Comparation
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Stream myStream;
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();

            OpenFileDialog1.Filter = "xlsx files (*.xlsx|*.xlsx|All files (*.*)|*.*";

            if (OpenFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if ((myStream = OpenFileDialog1.OpenFile()) != null)
                {
                    myStream.Close();
                }
                textBox1.Text = OpenFileDialog1.FileName;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Stream myStream;
            OpenFileDialog OpenFileDialog1 = new OpenFileDialog();

            OpenFileDialog1.Filter = "xlsx files (*.xlsx|*.xlsx|All files (*.*)|*.*";

            if (OpenFileDialog1.ShowDialog() == System.Windows.Forms.DialogResult.OK)
            {
                if ((myStream = OpenFileDialog1.OpenFile()) != null)
                {
                    myStream.Close();
                }
                textBox2.Text = OpenFileDialog1.FileName;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            FolderBrowserDialog FolderBrowserDialog1 = new FolderBrowserDialog();

            DialogResult result = FolderBrowserDialog1.ShowDialog();
            if (result == DialogResult.OK)
            {
                textBox3.Text = FolderBrowserDialog1.SelectedPath;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string FileName = textBox1.Text;
            object rOnly = true;
            object SaveChanges = false;
            object MissingObj = System.Reflection.Missing.Value;

            Excel.Application app = new Excel.Application();
            Excel.Workbooks workbooks = null;
            Excel.Workbook workbook = null;
            Excel.Sheets sheets = null;

            workbooks = app.Workbooks;
            workbook = workbooks.Open(FileName, MissingObj, rOnly, MissingObj, MissingObj,
                                        MissingObj, MissingObj, MissingObj, MissingObj, MissingObj,
                                        MissingObj, MissingObj, MissingObj, MissingObj, MissingObj);


            sheets = workbook.Sheets;

            Context db = new Context();

            toolStripStatusLabel1.Text = "Очистка таблицы 1 БД";
            var rows = from o in db.Table1
                       select o;
            foreach (var row in rows)
            {
                db.Table1.Remove(row);
            }
            db.SaveChanges();
            toolStripStatusLabel1.Text = "Очистка таблицы 1 БД завершена";

            toolStripStatusLabel1.Text = "Очистка таблицы 2 БД";
            var rows2 = from o in db.Table2
                        select o;
            foreach (var row2 in rows2)
            {
                db.Table2.Remove(row2);
            }
            db.SaveChanges();
            toolStripStatusLabel1.Text = "Очистка таблицы 2 БД завершена";

            toolStripStatusLabel1.Text = "Загрузка данных в таблицу 1 БД";
            foreach (Excel.Worksheet worksheet in sheets)
            {
                Excel.Range UsedRange = worksheet.UsedRange;
                Excel.Range urRows = UsedRange.Rows;

                int RowsCount = urRows.Count;
                int step = RowsCount / 50;
                int count = 0;
                for (int i = 1; i <= RowsCount; i++)
                {
                    Excel.Range CellRange = UsedRange.Cells[i, 1];

                    string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                        (CellRange as Excel.Range).Value2.ToString();

                    if (CellText != null)
                    {
                        var table1 = new Table1
                        {
                            Number1 = CellText
                        };

                        db.Table1.Add(table1);
                    }

                    toolStripStatusLabel1.Text = $"В таблицу 1 БД из {RowsCount} записей загружено: {i}";

                    count++;
                    if (count == step)
                    {
                        progressBar1.PerformStep();
                        count = 0;
                    }
                }

            }
            db.SaveChanges();
            toolStripStatusLabel1.Text = "Загрузка данных в таблицу 1 БД завершена";

            FileName = textBox2.Text;

            app = new Excel.Application();
            workbooks = null;
            workbook = null;
            sheets = null;

            workbooks = app.Workbooks;
            workbook = workbooks.Open(FileName, MissingObj, rOnly, MissingObj, MissingObj,
                                        MissingObj, MissingObj, MissingObj, MissingObj, MissingObj,
                                        MissingObj, MissingObj, MissingObj, MissingObj, MissingObj);


            sheets = workbook.Sheets;

            toolStripStatusLabel1.Text = "Загрузка данных в таблицу 2 БД";
            foreach (Excel.Worksheet worksheet in sheets)
            {
                Excel.Range UsedRange = worksheet.UsedRange;
                Excel.Range urRows = UsedRange.Rows;

                int RowsCount = urRows.Count;
                int step = RowsCount / 50;
                int count = 0;
                for (int i = 1; i <= RowsCount; i++)
                {
                    Excel.Range CellRange = UsedRange.Cells[i, 1];

                    string CellText = (CellRange == null || CellRange.Value2 == null) ? null :
                                        (CellRange as Excel.Range).Value2.ToString();

                    if (CellText != null)
                    {
                        var table2 = new Table2
                        {
                            Number2 = CellText
                        };

                        db.Table2.Add(table2);
                    }

                    toolStripStatusLabel1.Text = $"В таблицу 2 БД из {RowsCount} записей загружено: {i}";

                    count++;
                    if (count == step)
                    {
                        progressBar1.PerformStep();
                        count = 0;
                    }
                }

            }
            db.SaveChanges();
            toolStripStatusLabel1.Text = "Загрузка данных в таблицу 2 БД завершена";

            toolStripStatusLabel1.Text = "Выполняется запрос к БД";
            var result = from c in db.Table1
                         join p in db.Table2 on c.Number1 equals p.Number2 into ps
                         from p in ps.DefaultIfEmpty()
                         select new { Number1 = c, Number2 = p == null ? c.Number1 : null };
            toolStripStatusLabel1.Text = "Запрос к БД выполнен";

            toolStripStatusLabel1.Text = "Вывод в файл";
            StreamWriter f = new StreamWriter(textBox3.Text + @"\result.txt");
            foreach (var p in result)
            {
                if (p.Number2 != null)
                {
                    f.WriteLine(p.Number2);
                    Console.WriteLine(p.Number2);
                }
            }
            f.Close();
            toolStripStatusLabel1.Text = "Вывод в файл завершен";
            toolStripStatusLabel1.Text = "Готово";
        }
    }
}
