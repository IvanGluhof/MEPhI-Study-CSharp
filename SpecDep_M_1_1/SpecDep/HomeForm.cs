using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;
using System.Diagnostics;
using System.Data.OleDb;
using Excel = Microsoft.Office.Interop.Excel;
using Calendar.NET;
using System.Timers;

namespace SpecDep
{
    public partial class HomeForm : Form
    {
        private dynamic DB_frame;
        private dynamic DOC_frame;

        public HomeForm()
        {
            InitializeComponent();
            //1. Загружаем фреймворки
            InitializeFramework();
            //2. Устанавливаем внешний вид
            InitializeDesign();
            //3. Загружаем документы
            InitializeDataBoxes();
            //4. Загружаем календарь
            InitializeCalendar();
            //5. Инициируем таймер
            SetTimer();
        }

        #region Для конструктора
        /// <summary>
        /// Загрузка классов для работы с БД и документами.
        /// </summary>
        private void InitializeFramework()
        {
            DB_frame = new DataBase_FrameWork(this); // Инициализируем класс для работы с базой данных. Передаем форму для работы
            DOC_frame = new Documents_Framework(this); // Инициализируем класс для работы с документами
        }

        /// <summary>
        /// Устанавливает внешний вид TabControl
        /// </summary>
        private void InitializeDesign()
        {
            tabControl1.Appearance = TabAppearance.FlatButtons; tabControl1.ItemSize = new Size(0, 1); tabControl1.SizeMode = TabSizeMode.Fixed;
        }

        /// <summary>
        /// Подгрузка документов в textBox'ы
        /// </summary>
        protected internal void InitializeDataBoxes()
        {
            textBox_oshibki.Text = DOC_frame.Load_Oshibki();
            {
                if (textBox_oshibki.Text != "0") { this.textBox_oshibki.ForeColor = Color.Firebrick; }
                else { textBox_oshibki.ForeColor = Color.YellowGreen; }
            }
            textBox_sogl.Text = DB_frame.Load_Sogl();
            {
                if (textBox_sogl.Text != "0") { this.textBox_sogl.ForeColor = Color.Firebrick; }
                else { textBox_sogl.ForeColor = Color.YellowGreen; }
            }
            textBox_plat_por.Text = DB_frame.Load_plat_por();
            {
                if (textBox_plat_por.Text != "0") { this.textBox_plat_por.ForeColor = Color.Firebrick; }
                else { textBox_plat_por.ForeColor = Color.YellowGreen; }
            }
            textBox_correct_otchet.Text = DB_frame.Load_osh_CB();
            {
                if (textBox_correct_otchet.Text != "0") { this.textBox_correct_otchet.ForeColor = Color.Firebrick; }
                else { textBox_correct_otchet.ForeColor = Color.YellowGreen; }
            }
            textBox_make_otchet.Text = DOC_frame.Load_kol_make_otch();
        }

        /// <summary>
        /// Настройка календаря
        /// </summary>
        private void InitializeCalendar()
        {
            DB_frame.Read_Dates_For_Calendar();
            calendar1.LoadPresetHolidays = false;
            calendar1.AllowEditingEvents = true;
            calendar1.today();
            var Today_Event = new CustomEvent
            {
                Date = DateTime.Today.Date,
                EventText = "Сформировать отчет" + Environment.NewLine + "на сегодня",
                EventColor = Color.SkyBlue
            };
            this.calendar1.AddEvent(Today_Event);

            var Yesterday_Event = new CustomEvent
            {
                Date = DateTime.Today.Date.AddDays(-1),
                EventText = "Сформировать отчет",
                EventColor = Color.SkyBlue
            };
            this.calendar1.AddEvent(Yesterday_Event);
        }
        #endregion

        private void calendar1_MouseEnter(object sender, EventArgs e)
        {
            label1.Text = this.calendar1.Selected_date;
        }

        private void открытьToolStripMenuItem_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.Title = "Выберите файл";
            ofd.InitialDirectory = new DirectoryInfo(@".\\Debug").FullName;
            if (ofd.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            Process.Start(ofd.FileName);
        }
    }     
}
