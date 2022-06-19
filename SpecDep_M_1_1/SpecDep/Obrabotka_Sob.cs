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
using System.Timers;
using Calendar.NET;

namespace SpecDep
{
    public partial class HomeForm : Form
    {
        List<CustomEvent> events_list = new List<CustomEvent>();
        /// <summary>
        /// Функция обработки события - возврата на главную страницу
        /// По возвращении на страницу происходит обновление
        /// </summary>
        private void tabpage_main_Enter(object sender, EventArgs e)
        {
            InitializeDataBoxes();
        }

        #region Всё, что с текст боксами
        // Клики
        private void textBox_oshibki_MouseClick(object sender, MouseEventArgs e)
        {
            this.tabControl1.SelectedTab = tabpage_oshibki;
        }
        private void textBox_sogl_MouseClick(object sender, MouseEventArgs e)
        {
            this.tabControl1.SelectedTab = tabpage_work_with_BD;
        }
        private void textBox_plat_por_MouseClick(object sender, MouseEventArgs e)
        {
            this.tabControl1.SelectedTab = tabpage_work_with_BD;
        }
        private void textBox_make_otchet_MouseClick(object sender, MouseEventArgs e)
        {
            this.tabControl1.SelectedTab = tabPage_make_otchet;
        }
        private void textBox_correct_otchet_MouseClick(object sender, MouseEventArgs e)
        {
            this.tabControl1.SelectedTab = tabPage_make_otchet;
            textBox_correct_otchet_2.Text = DB_frame.Load_osh_CB();
            {
                if (textBox_correct_otchet_2.Text != "0") { this.textBox_correct_otchet_2.ForeColor = Color.Firebrick; }
                else { textBox_correct_otchet_2.ForeColor = Color.YellowGreen; }
            }
        }
        private void textBox_go_to_spravoch_MouseClick(object sender, MouseEventArgs e)
        {
            this.tabControl1.SelectedTab = tabPage_work_with_spravoch;
        }

        // Выделение
        private void textBox_oshibki_MouseEnter(object sender, EventArgs e)
        {
            textBox_oshibki.BorderStyle = BorderStyle.FixedSingle;
        }

        private void textBox_oshibki_MouseLeave(object sender, EventArgs e)
        {
            textBox_oshibki.BorderStyle = BorderStyle.None;
        }

        private void textBox_go_to_spravoch_MouseEnter(object sender, EventArgs e)
        {
            textBox_go_to_spravoch.BorderStyle = BorderStyle.FixedSingle;
        }

        private void textBox_go_to_spravoch_MouseLeave(object sender, EventArgs e)
        {
            textBox_go_to_spravoch.BorderStyle = BorderStyle.None;
        }

        private void textBox_sogl_MouseEnter(object sender, EventArgs e)
        {
            textBox_sogl.BorderStyle = BorderStyle.FixedSingle;
        }

        private void textBox_sogl_MouseLeave(object sender, EventArgs e)
        {
            textBox_sogl.BorderStyle = BorderStyle.None;
        }

        private void textBox_plat_por_MouseEnter(object sender, EventArgs e)
        {
            textBox_plat_por.BorderStyle = BorderStyle.FixedSingle;
        }

        private void textBox_plat_por_MouseLeave(object sender, EventArgs e)
        {
            textBox_plat_por.BorderStyle = BorderStyle.None;
        }

        private void textBox_make_otchet_MouseEnter(object sender, EventArgs e)
        {
            textBox_make_otchet.BorderStyle = BorderStyle.FixedSingle;
        }

        private void textBox_make_otchet_MouseLeave(object sender, EventArgs e)
        {
            textBox_make_otchet.BorderStyle = BorderStyle.None;
        }

        private void textBox_correct_otchet_MouseEnter(object sender, EventArgs e)
        {
            textBox_correct_otchet.BorderStyle = BorderStyle.FixedSingle;
        }

        private void textBox_correct_otchet_MouseLeave(object sender, EventArgs e)
        {
            textBox_correct_otchet.BorderStyle = BorderStyle.None;
        }

        #endregion

        #region фоновая загрузка + надпись
        protected internal System.Timers.Timer aTimer;
        protected internal void SetTimer()
        {
            aTimer = new System.Timers.Timer(40000); //1 sec = 1000 msec
            aTimer.Elapsed += OnTimedEvent;
            aTimer.AutoReset = true;
            aTimer.Enabled = true;
        }
        protected internal void OnTimedEvent(Object source, ElapsedEventArgs e)
        {
            backgroundWorker_download_doc.RunWorkerAsync();
        }
        protected internal void backgroundWorker_download_doc_DoWork(object sender, DoWorkEventArgs e)
        {
            DOC_frame.Load_doc_cycle();
        }

        protected internal void backgroundWorker_download_doc_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            //InitializeDataBoxes();
            SetText("Последняя загрузка:" + Environment.NewLine + Environment.NewLine + DateTime.Now);
        }

        private void SetText(string download_time)
        {
            if (InvokeRequired)
            {
                Invoke(new Action<string>((s) => textBox_time_refresh.Text = s), download_time);
                Invoke(new Action(InitializeDataBoxes));
            }
            else
            {
                textBox_time_refresh.Text = download_time;
                InitializeDataBoxes();
            }
        }

        protected internal void textBox_time_refresh_TextChanged(object sender, EventArgs e)
        {
            this.textBox_time_refresh.ForeColor = Color.DarkGreen;
        }
        #endregion

        #region Страница для работы с БД
        private void comboBox_view_SelectionChangeCommitted(object sender, EventArgs e)
        {
            DB_frame.View_Table();
        }

        private void toolStripButton_view_Click(object sender, EventArgs e)
        {
            DB_frame.View_Table();
        }
        private void toolStripButton_apply_filter_Click(object sender, EventArgs e)
        {
            int try_date = 0;
            if (!string.IsNullOrWhiteSpace(textBox_Data.Text))
                try
                {
                    DateTime data_vvod = DateTime.Parse(textBox_Data.Text);
                    try_date = 1;
                }
                catch
                {
                    MessageBox.Show("       Неверный формат даты!" + Environment.NewLine + Environment.NewLine + "  Введите дату в формате:" + "    дд.мм.гггг", "Внимание!");
                }
            else { try_date = 1; }
            if (try_date == 1)
            { DB_frame.Do_Filter(); }
        }
        private void toolStripButton_clear_filter_Click(object sender, EventArgs e)
        {
            textBox_Data.Clear();
            textBox_Naznach.Clear();
            DB_frame.View_Table();
        }
        private void toolStripButton_make_otkaz_sogl_plat_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd4 = new OpenFileDialog();
            ofd4.DefaultExt = "*.xls;*.xlsx";
            ofd4.Filter = "Microsoft Excel (*.xls*)|*.xls*";
            ofd4.Title = "Выберите файл, для создания отказа";
            ofd4.InitialDirectory = new DirectoryInfo(@".\\Архив").FullName;       //F:\Учеба\Программа\Ошибки

            if (ofd4.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            string file_name_obraz_proved = Path.GetDirectoryName(ofd4.FileName) + "\\" + "Образец отказа на проведение операции";
            string file_name_obraz = Path.GetDirectoryName(ofd4.FileName) + "\\" + "Образец отказа оформления";
            if (ofd4.FileName.Contains("плат_пор"))
            {
                Excel.Sheets excelsheets;
                Excel.Range excelcells;
                Excel.Application objWorkExcel = new Excel.Application();

                //откроем данный файл
                Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(ofd4.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                excelsheets = objWorkBook.Worksheets;                                       //Получаем массив ссылок на листы выбранной книги
                Excel.Worksheet excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);  //Получаем ссылку на лист 1
                //вытащим название УК
                Excel.Range CellUK = excelworksheet.get_Range("A12");
                string nameUK = CellUK.Text.ToString();
                string name_file_with_osh = Path.GetFileNameWithoutExtension(ofd4.FileName);
                string data = Convert.ToString(DateTime.Now);

                //откроем образец файла отказа
                Excel.Workbook objWorkBook_2 = objWorkExcel.Workbooks.Open(file_name_obraz, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                excelsheets = objWorkBook_2.Worksheets;                                       //Получаем массив ссылок на листы выбранной книги
                Excel.Worksheet excelworksheet_2 = (Excel.Worksheet)excelsheets.get_Item(1);  //Получаем ссылку на лист 1
                //запишем все в файл
                excelcells = excelworksheet_2.get_Range("H12", Type.Missing);
                excelcells.Value2 = data.Substring(0, 10);

                excelcells = excelworksheet_2.get_Range("B15", Type.Missing);
                excelcells.Value2 = nameUK;

                excelcells = excelworksheet_2.get_Range("B18", Type.Missing);
                excelcells.Value2 = name_file_with_osh;

                string name_otkaz_file = Path.GetDirectoryName(ofd4.FileName) + "\\" + "Отказ на " + name_file_with_osh + ".xls";
                objWorkExcel.DisplayAlerts = false;
                objWorkBook_2.SaveAs(name_otkaz_file, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                objWorkBook_2.Close();
                objWorkBook.Close();
                objWorkExcel.Quit();

                string putv_for_all = new DirectoryInfo(@".\\Отчеты_на_отправку").FullName;             //путь куда 
                string putiz_file = ofd4.FileName;
                string putv_file = putv_for_all + "\\" + Path.GetFileNameWithoutExtension(ofd4.FileName) + ".xlsx";
                File.Move(putiz_file, putv_file);

                string new_name_otkaz_file = putv_for_all + "\\" + "Отказ на " + name_file_with_osh + ".xls";
                File.Move(name_otkaz_file, new_name_otkaz_file);                                //перемещение

                MessageBox.Show("Уведомление об отказе создано и отправлено!" + "\r\n" + "\r\n" + "Необходимо удалить соответствующую строку в базе данных!", "Уведомление");
            }

            if (ofd4.FileName.Contains("согласие"))
            {
                Excel.Sheets excelsheets;
                Excel.Range excelcells;
                Excel.Application objWorkExcel = new Excel.Application();

                //откроем данный файл
                Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(ofd4.FileName, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                excelsheets = objWorkBook.Worksheets;                                       //Получаем массив ссылок на листы выбранной книги
                Excel.Worksheet excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);  //Получаем ссылку на лист 1
                //вытащим название УК
                Excel.Range CellUK = excelworksheet.get_Range("F2");
                string nameUK = CellUK.Text.ToString();
                string name_file_with_osh = Path.GetFileNameWithoutExtension(ofd4.FileName);
                string data = Convert.ToString(DateTime.Now);

                //откроем образец файла отказа
                Excel.Workbook objWorkBook_2 = objWorkExcel.Workbooks.Open(file_name_obraz_proved, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                excelsheets = objWorkBook_2.Worksheets;                                       //Получаем массив ссылок на листы выбранной книги
                Excel.Worksheet excelworksheet_2 = (Excel.Worksheet)excelsheets.get_Item(1);  //Получаем ссылку на лист 1
                //запишем все в файл
                excelcells = excelworksheet_2.get_Range("H12", Type.Missing);
                excelcells.Value2 = data.Substring(0, 10);

                excelcells = excelworksheet_2.get_Range("B15", Type.Missing);
                excelcells.Value2 = nameUK;

                excelcells = excelworksheet_2.get_Range("B18", Type.Missing);
                excelcells.Value2 = name_file_with_osh;

                string name_otkaz_file = Path.GetDirectoryName(ofd4.FileName) + "\\" + "Отказ на проведение " + name_file_with_osh + ".xls";
                objWorkExcel.DisplayAlerts = false;
                objWorkBook_2.SaveAs(name_otkaz_file, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                objWorkBook_2.Close();
                objWorkBook.Close();
                objWorkExcel.Quit();

                string putv_for_all = new DirectoryInfo(@".\\Отчеты_на_отправку").FullName;             //путь куда 
                string putiz_file = ofd4.FileName;
                string putv_file = putv_for_all + "\\" + Path.GetFileNameWithoutExtension(ofd4.FileName) + ".xlsx";
                File.Move(putiz_file, putv_file);

                string new_name_otkaz_file = putv_for_all + "\\" + "Отказ на проведение " + name_file_with_osh + ".xls";
                File.Move(name_otkaz_file, new_name_otkaz_file);

                MessageBox.Show("Уведомление об отказе создано и отправлено!" + "\r\n" + "\r\n" + "Необходимо удалить соответствующую строку в базе данных!", "Уведомление");
            }
        }

        private void Save_toolStripButton_Click(object sender, EventArgs e)
        {
            DB_frame.Save_to_BD();
        }

        private void Delete_toolStripButton_Click(object sender, EventArgs e)
        {
            DB_frame.Delete_Save_Row();
        }

        #region Календарь выпадающий
        private void textBox_Data_TextChanged(object sender, EventArgs e)
        {
            monthCalendar2.Show();
        }

        private void textBox_Data_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            monthCalendar2.Show();
        }

        private void monthCalendar2_MouseLeave(object sender, EventArgs e)
        {
            monthCalendar2.Hide();
        }        
        
        private void monthCalendar2_DateSelected(object sender, DateRangeEventArgs e)
        {
            textBox_Data.Text = monthCalendar2.SelectionStart.ToShortDateString().ToString();
            monthCalendar2.Hide();
        }
        #endregion
        #endregion

        #region Страница для работы с ошибками
        private void tabpage_oshibki_Enter(object sender, EventArgs e)
        {
            Form_Update_list_oshibki();
        }

        private void toolStripButton_to_work_with_spravoch_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedTab = tabPage_work_with_spravoch;
        }

        private void toolStripButton_success_osh_Click(object sender, EventArgs e)
        {
            DOC_frame.Load_success_osh(name_file_selected_at_listbox);
        }

        private void toolStripButton_make_otkaz_Click(object sender, EventArgs e)
        {
            DOC_frame.Load_make_otkaz(name_file_selected_at_listbox);
        }

        private void toolStripButton_opis_osh_Click(object sender, EventArgs e)
        {
            DOC_frame.Load_file_osh(name_file_selected_at_listbox);
            textBox_vybran_file.Text = name_file_selected_at_listbox;
        }

        private void toolStripButton_refresh_Click(object sender, EventArgs e)
        {
            Form_Update_list_oshibki();
        }
        private void toolStripButton_open_directory_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls";
            ofd.Filter = "Текстовые файлы (*.xls*)|*.xls*";
            ofd.InitialDirectory = new DirectoryInfo(@".\\Ошибки").FullName;
            ofd.ShowDialog();
            Process.Start(ofd.FileName);
        }

        #region ListBox и его ОООООЧЕНЬ тонкая настройка
        private void Form_Update_list_oshibki()
        {
            listBox_oshibki.DrawMode = DrawMode.Normal;
            listBox_oshibki.Items.Clear();          

            List<string> files = DOC_frame.Count_Oshibki_to_ListBox();
            List<string[]> file_oshibki = DOC_frame.Get_Oshibki;

            for (int i = 0; i < files.Count; i++)
            {
                string listbox_oshibka = files[i] + "\n";
                for (int j = 0; j < file_oshibki[i].Length; j++)
                {
                    listbox_oshibka = listbox_oshibka + "Текст ошибки: " + file_oshibki[i][j] + "\n";
                }
                listbox_oshibka = listbox_oshibka + "\n";
                listBox_oshibki.Items.Add(listbox_oshibka);
            }
            textBox_oshibki_2.Text = DOC_frame.Load_Oshibki();
            {
                if (textBox_oshibki_2.Text != "0") { this.textBox_oshibki_2.ForeColor = Color.Firebrick; }
                else { textBox_oshibki_2.ForeColor = Color.YellowGreen; }
            }

            listBox_oshibki.DrawMode = DrawMode.OwnerDrawVariable;
        }

        private int ItemMargin = 10;
        private void listBox_oshibki_MeasureItem(object sender, MeasureItemEventArgs e)
        {
            // Наши ListBox и его ListBoxItem
            ListBox lst = sender as ListBox;
            string txt = (string)lst.Items[e.Index];

            // Строка
            SizeF txt_size = e.Graphics.MeasureString(txt, this.Font);

            // Установка размера
            e.ItemHeight = (int)txt_size.Height + 4 * ItemMargin;
            e.ItemWidth = (int)txt_size.Width;
        }
        private void listBox_oshibki_DrawItem(object sender, DrawItemEventArgs e)
        {
            // Наши ListBox и его ListBoxItem
            ListBox lst = sender as ListBox;
            string txt = (string)lst.Items[e.Index];

            // Задник
            e.DrawBackground();

            // Проверка, если выбран ListBoxItem
            if ((e.State & DrawItemState.Selected) == DrawItemState.Selected)
            {
                // = да -> подсвечиваем
                e.Graphics.DrawString(txt, new Font("Times New Roman", 12), SystemBrushes.HighlightText, e.Bounds.Left, e.Bounds.Top + ItemMargin);
            }
            else
            {
                // = нет -> не подствечиаем. бэк стандартным
                using (SolidBrush br = new SolidBrush(e.ForeColor))
                {
                    e.Graphics.DrawString(txt, new Font("Times New Roman", 12), br, e.Bounds.Left, e.Bounds.Top + ItemMargin);
                }
            }

            //Линии
            using (Graphics g = e.Graphics)
            {
                g.DrawRectangle(new Pen(Color.Black), e.Bounds);
            }

            // Отрисовка фокуса
            e.DrawFocusRectangle();
        }

        string name_file_selected_at_listbox; //переменная с именем
        private void listBox_oshibki_SelectedIndexChanged(object sender, EventArgs e)
        {
            try
            {
                var selected_item = listBox_oshibki.SelectedItem.ToString();
                name_file_selected_at_listbox = DOC_frame.file_name[listBox_oshibki.SelectedIndex].ToString();
            }
            catch
            {
                MessageBox.Show("Ничего не выбрано");
            }
        }
        #endregion

        #endregion

        #region Страница для работы со справочниками
        private void toolStripButton_sp_Click(object sender, EventArgs e)
        {
            DB_frame.View_Table_sp();
        }
        private void comboBox_sp_SelectionChangeCommitted(object sender, EventArgs e)
        {
            DB_frame.View_Table_sp();
        }
        private void toolStripButton_del_sp_Click(object sender, EventArgs e)
        {
            DB_frame.Delete_Save_Row_sp();
        }

        private void toolStripButton_save_sp_Click(object sender, EventArgs e)
        {
            DB_frame.Save_to_BD_sp();
        }
        #endregion

        #region Кнопка назад для всех вкладок + событие возврата
        private void toolStripButton1_back_Click(object sender, EventArgs e)
        {
            return_home();
        }

        private void toolStripButton_back2_Click(object sender, EventArgs e)
        {
            return_home();
        }

        private void toolStripButton_back3_Click(object sender, EventArgs e)
        {
            return_home();
        }

        private void toolStripButton_back4_Click(object sender, EventArgs e)
        {
            return_home();
        }

        private void toolStripButton_back_to_osh_Click(object sender, EventArgs e)
        {
            this.tabControl1.SelectedTab = tabpage_oshibki;
        }

        private void return_home()
        {
            this.tabControl1.SelectedTab = tabpage_main;
        }
        #endregion

        #region Выход и сворачивание в трей
        private void HomeForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно хотите выйти из приложения?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                Environment.Exit(0);
            }
            else
            {
                e.Cancel = true;
            }

        }
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно хотите выйти из приложения?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                Environment.Exit(0);
            }
        }

        FormWindowState OldFormState = new FormWindowState();
        private void HomeForm_Resize(object sender, EventArgs e)
        {
            OldFormState = WindowState;
            if (FormWindowState.Minimized == WindowState)
            {
                notifyIcon1.Visible = true;
                Hide();
            }
        }
        private void notifyIcon1_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            Show();
            WindowState = OldFormState;
        }
        #endregion

        #region Работа с отчетами
        private void toolStripButton_refresh_kalend_Click(object sender, EventArgs e)
        {
            if (comboBox_UK_otch.SelectedItem == null || comboBox_Fond_otch.SelectedItem == null)
            {
                MessageBox.Show("Одно или несколько из полей не выбраны");
            }
            else
            {
                if (events_list.Count != 0)
                {
                    foreach (CustomEvent _event in events_list)
                    {
                        this.calendar1.RemoveEvent(_event);
                    }
                    events_list.Clear();
                }
                DB_frame.Refresh_Dates_For_Calendar(comboBox_UK_otch.SelectedItem.ToString(), comboBox_Fond_otch.SelectedItem.ToString());
                for (int i = 0; i < DB_frame.Dates.Count; i++)
                {
                    var DB_Date_Event = new CustomEvent
                    {
                        Date = DateTime.Parse(DB_frame.Dates[i]),
                        EventText = DB_frame.Values[i],
                        EventColor = Color.SkyBlue
                    };                 
                    events_list.Add(DB_Date_Event);
                }

                foreach(CustomEvent _event in events_list)
                {
                    this.calendar1.AddEvent(_event);
                }
                textBox_correct_otchet_2.Text = DB_frame.Load_osh_CB();
                {
                    if (textBox_correct_otchet_2.Text != "0") { this.textBox_correct_otchet_2.ForeColor = Color.Firebrick; }
                    else { textBox_correct_otchet_2.ForeColor = Color.YellowGreen; }
                }
            }
        }
        private void toolStripButton_make_otchet_Click(object sender, EventArgs e)      //кнопка создания отчета
        {
            DOC_frame.Make_otchet();
        }

        private void toolStripButton_see_otchet_Click(object sender, EventArgs e)
        {
            DOC_frame.Watch_otchet();
        }
        private void toolStripButton_open_uved_Click(object sender, EventArgs e)
        {
            OpenFileDialog ofd2 = new OpenFileDialog();
            ofd2.DefaultExt = "*.xls;*.xlsx";
            ofd2.FileName = "ошибки_*";
            ofd2.Filter = "Microsoft Excel (*.xls*)|*.xls";
            ofd2.Title = "Выберите файл уведомления для просмотра ошибок в отчете";
            ofd2.InitialDirectory = new DirectoryInfo(@".\\Архив_ошибок_в_отчетах").FullName;                    //сразу путь в папку ошибок F:\Учеба\Программа\Ошибки

            if (ofd2.ShowDialog() != DialogResult.OK)
            {
                return;
            }
            MessageBox.Show("После исправления отчета необходимо проставить отметки в БД, а также удалить уведомление от ЦБ!", "Уведомление");
            Process.Start(ofd2.FileName);
        }
        #endregion
    }
}