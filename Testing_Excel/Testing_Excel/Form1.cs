using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
//для Excel
using System.Reflection;
using ExcelObj = Microsoft.Office.Interop.Excel;

namespace Testing_Excel
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            //Задаем расширение имени файла по умолчанию.
            ofd.DefaultExt = "*.xls;*.xlsx";
            //Задаем строку фильтра имен файлов, которая определяет
            //варианты, доступные в поле "Файлы типа" диалогового
            //окна.
            ofd.Filter = "Excel Sheet(*.xlsx)|*.xlsx";
            //Задаем заголовок диалогового окна.
            ofd.Title = "Выберите документ для загрузки данных";
            ExcelObj.Application app = new ExcelObj.Application();
            ExcelObj.Workbook workbook;
            ExcelObj.Worksheet NewSheet;
            ExcelObj.Range SheetRange;
            DataTable dt = new DataTable();

            if (ofd.ShowDialog() == DialogResult.OK)
            {
                label1.Text = ofd.FileName;

                workbook = app.Workbooks.Open(ofd.FileName, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value, Missing.Value, Missing.Value, Missing.Value,
                Missing.Value);

                NewSheet = (ExcelObj.Worksheet)workbook.Sheets.get_Item(1);
                SheetRange = NewSheet.UsedRange;

                for (int ClmNum = 1; ClmNum <= SheetRange.Columns.Count; ClmNum++)
                {
                    dt.Columns.Add
                        (new DataColumn((SheetRange.Cells[1, ClmNum] as ExcelObj.Range).Value2.ToString()));
                }

                for (int Rnum = 2; Rnum <= SheetRange.Rows.Count; Rnum++)
                {
                    DataRow dr = dt.NewRow();
                    for (int Cnum = 1; Cnum <= SheetRange.Columns.Count; Cnum++)
                    {
                        if ((SheetRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2 != null)
                        {
                            dr[Cnum - 1] =
                            (SheetRange.Cells[Rnum, Cnum] as ExcelObj.Range).Value2.ToString();
                        }
                    }
                    dt.Rows.Add(dr);
                    dt.AcceptChanges();
                }
                dataGridView1.DataSource = dt;
                app.Quit();
            }
            else
            {
                Application.Exit();
            }

        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            // Создаём переменную rsl, которая будет хранить результат вывода окна с вопросом
            // (пользователь нажал одну из клавиш на окне - это и есть результат)
            // MessageBox будет содержать вопрос, а также кнопки Yes, No и иконку Question
            DialogResult rsl = MessageBox.Show("Вы действительно хотите выйти из приложение?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Question);

            // если пользователь нажал кнопку да
            if (rsl == DialogResult.Yes)
            {
                // выходим из приложение
                Application.Exit();
            
            }
        }

        private void panel1_Paint(object sender, PaintEventArgs e)
        {
            float W = panel1.Width, H = panel1.Height;
            float halfW = W / 2, halfH = H / 2;
            // оси координат
            e.Graphics.DrawLine(Pens.Black, halfW, 0, halfW, H);
            e.Graphics.DrawLine(Pens.Black, 0, halfH, W, halfH);
            /*/ координаты предыдущей точки
            int ixPrev = -1, iyPrev = (int)halfH;
            // тангенс на интервале x=[-Pi..Pi]
            // проходим по всем точкам на форме, вычисляем x и y=tg(x)
            for (int ix = 0; ix < W; ix++)
            {
                // переводим x в диапазон -1..1
                float x = (ix - halfW) / halfW;
                // переводим x в -pi..pi
                x *= (float)Math.PI;

                // получаем tg(x)
                float y = (float)Math.Tan(x);
                // переводим y из -1..1 в пикселы на форме
                int iy = (int)(halfH - y * halfH);
                // вуаля
                e.Graphics.DrawLine(Pens.Red, ixPrev, iyPrev, ix, iy);
                ixPrev = ix;
                iyPrev = iy;
            }
             */ 
        }
    }
}
