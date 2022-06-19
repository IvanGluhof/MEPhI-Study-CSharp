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
using System.Reflection;
using Excel = Microsoft.Office.Interop.Excel;


namespace SpecDep
{
    /// <summary>
    /// Класс для работы только с физическими документами.
    /// т.е. с теми, которые лежат только в папках
    /// </summary>
    class Documents_Framework
    {
        private HomeForm work_form;

        string filename;
        string fullnametxt;

        static string path_to_oshibki;
        static string path_to_zagruzki;
        static string path_to_arh_otch;

        public dynamic file_names;
        public dynamic file_name
        {
            get { return file_names; }
        }

        public Documents_Framework(HomeForm form)
        {
            work_form = form;
            Set_Paths();
        }
        
        private void Set_Paths()
        {
            path_to_oshibki = (Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Ошибки");
            path_to_zagruzki = (Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Загрузка");
            path_to_arh_otch = (Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) + "\\Архив_отчетов");
        }

        /// <summary>
        /// Будет показывать сколько файлов с ошибками на данный момент
        /// </summary>
        /// <returns></returns>
        protected internal string Load_Oshibki()
        {
            var colfiles = new DirectoryInfo(path_to_oshibki).GetFiles().Length.ToString();    //F:\Учеба\Программа\Ошибки
            int kolichfilov = Convert.ToInt32(colfiles);
            string result = Convert.ToString((kolichfilov / 2)-1);
            return result;
        }

        /*не используется
        protected internal string Load_opis_oshibok()
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.txt";
            ofd.Filter = "Текстовые файлы (*.txt*)|*.txt*";
            ofd.Title = "Выберите файл ошибок для исправления";
            ofd.InitialDirectory = new DirectoryInfo(path_to_oshibki).FullName;         //сразу путь в папку ошибок F:\Учеба\Программа\Ошибки
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Process.Start(ofd.FileName);                                        //запускаем Блокнот
                filename = Path.GetFileNameWithoutExtension(ofd.FileName);          //отрезаем расширение от текстового файла
                fullnametxt = ofd.FileName;                                         //присваиваем путь переменной для удаления файла в конце исправления               

                string file_put = Path.GetDirectoryName(ofd.FileName);                  //получаем директорию
                string full_file_name_excel = file_put + "\\" + filename + ".xlsx";     //получаем путь + имя файла в формате excel
                Process.Start(full_file_name_excel);                                    //запускаем excel
                return filename;
            }
            return "none";
        }*/

        /// <summary>
        /// показывает количество отчетов к составлению
        /// </summary>
        protected internal string Load_kol_make_otch()
        {
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DataBaseMy.accdb";
            using (OleDbConnection conn = new OleDbConnection(connectionString)) //новое подключение к БД
            {
                OleDbCommand CommandBD = new OleDbCommand();                                      //команда, через которую все делается
                CommandBD.Connection = conn;                                                      //соединение с бд
                conn.Open();
                CommandBD.CommandText = "SELECT count(*) FROM Список_фондов";
                int kol_strok = 2 * (Convert.ToInt32(CommandBD.ExecuteScalar())); //подсчет количества строк (х2 т.к. выписка и СЧА)
                int kol_fondov = kol_strok/2;                                       // /2 т.к. выше есть х2
                string name_of_fond;
                DateTime data_today_full = DateTime.Now;
                string data_now = data_today_full.ToShortDateString();
                var obrazec_directory = new DirectoryInfo(@".\\Архив_отчетов");

                CommandBD.CommandText = "SELECT * FROM Список_фондов";
                OleDbDataReader dr1 = CommandBD.ExecuteReader();

                foreach (FileInfo file in obrazec_directory.GetFiles())
                {
                    while (dr1.Read())
                    {
                        name_of_fond = dr1.GetString(2);
                        if ((file.Name.Contains(name_of_fond)) && (file.Name.Contains(data_now)))
                        {
                            kol_strok = kol_strok - 1;
                        }
                    }
                }
                string result = Convert.ToString(kol_strok + " по " + kol_fondov + " фондам");
                return result;
            }
        }
        
        protected internal void Load_file_osh(string nazv)
        {
            string f = path_to_oshibki + "\\" + nazv;
            Process.Start(f);
        }

        protected internal void Load_success_osh(string nazv_2)
        {
            string f = path_to_oshibki + "\\" + nazv_2;

            string fi = Path.GetFileNameWithoutExtension(f);         //отрезает расширение от файла
            string fileName2 = fi + ".xlsx";                                       //подписываем новое расширение
            string fullnametxt = fi + ".txt";
            string putiz = new DirectoryInfo(path_to_oshibki).FullName;                        //путь откуда F:\Учеба\Программа\Ошибки
            string putv = new DirectoryInfo(path_to_zagruzki).FullName;                       //путь куда F:\Учеба\Программа\Загрузка
            string sourceFile = Path.Combine(putiz, nazv_2);
            string destFile = Path.Combine(putv, fileName2);
            string destFiletxt = Path.Combine(putiz, fullnametxt);

            File.Move(sourceFile, destFile);        //перемещение
            File.Delete(destFiletxt);               //Удаляет текстовый файл
            MessageBox.Show("Успешно перемещен в 'Загрузки'");
        }

        protected internal void Load_make_otkaz(string nazv_3)
        {
            string f = path_to_oshibki + "\\" + nazv_3;
            string file_name_obraz = path_to_oshibki + "\\" + "Образец отказа";
            if (nazv_3.Contains("платежн"))
            {
                Excel.Sheets excelsheets;
                Excel.Range excelcells;
                Excel.Application objWorkExcel = new Excel.Application();

                //откроем данный файл
                Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(f, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                excelsheets = objWorkBook.Worksheets;                                       //Получаем массив ссылок на листы выбранной книги
                Excel.Worksheet excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);  //Получаем ссылку на лист 1
                //вытащим название УК
                Excel.Range CellUK = excelworksheet.get_Range("A12");
                string nameUK = CellUK.Text.ToString();
                string name_file_with_osh = Path.GetFileNameWithoutExtension(f);
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

                string name_otkaz_file = Path.GetDirectoryName(path_to_oshibki) + "\\" + "Отказ на " + name_file_with_osh + ".xls";
                objWorkExcel.DisplayAlerts = false;
                objWorkBook_2.SaveAs(name_otkaz_file, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                objWorkBook_2.Close();
                objWorkBook.Close();
                objWorkExcel.Quit();

                string putv_for_all = new DirectoryInfo(@".\\Отчеты_на_отправку").FullName;             //путь куда 
                string putiz_file = f;
                string putv_file = putv_for_all + "\\" + Path.GetFileNameWithoutExtension(f) + ".xlsx";
                File.Move(putiz_file, putv_file);

                string new_name_otkaz_file = putv_for_all + "\\" + "Отказ на " + name_file_with_osh + ".xls";
                File.Move(name_otkaz_file, new_name_otkaz_file);                                //перемещение
                string name_file_txt = path_to_oshibki + "\\" + Path.GetFileNameWithoutExtension(f) + ".txt";
                File.Delete(name_file_txt);

                MessageBox.Show("Уведомление об отказе создано и отправлено!");
            }

            if (nazv_3.Contains("согласи"))
            {
                Excel.Sheets excelsheets;
                Excel.Range excelcells;
                Excel.Application objWorkExcel = new Excel.Application();

                //откроем данный файл
                Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(f, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                excelsheets = objWorkBook.Worksheets;                                       //Получаем массив ссылок на листы выбранной книги
                Excel.Worksheet excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);  //Получаем ссылку на лист 1
                //вытащим название УК
                Excel.Range CellUK = excelworksheet.get_Range("F2");
                string nameUK = CellUK.Text.ToString();
                string name_file_with_osh = Path.GetFileNameWithoutExtension(f);
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

                string name_otkaz_file = Path.GetDirectoryName(path_to_oshibki) + "\\" + "Отказ на " + name_file_with_osh + ".xls";
                objWorkExcel.DisplayAlerts = false;
                objWorkBook_2.SaveAs(name_otkaz_file, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                objWorkBook_2.Close();
                objWorkBook.Close();
                objWorkExcel.Quit();

                string putv_for_all = new DirectoryInfo(@".\\Отчеты_на_отправку").FullName;             //путь куда 
                string putiz_file = f;
                string putv_file = putv_for_all + "\\" + Path.GetFileNameWithoutExtension(f) + ".xlsx";
                File.Move(putiz_file, putv_file);

                string new_name_otkaz_file = putv_for_all + "\\" + "Отказ на " + name_file_with_osh + ".xls";
                File.Move(name_otkaz_file, new_name_otkaz_file);
                string name_file_txt = path_to_oshibki + "\\" + Path.GetFileNameWithoutExtension(f) + ".txt";
                File.Delete(name_file_txt);

                MessageBox.Show("Уведомление об отказе создано и отправлено!");
            }
            if ((!nazv_3.Contains("платежн")) && (!nazv_3.Contains("согласи")))
            {
                MessageBox.Show("Неизвестный файл, сформируйте отказ в ручную!");
            }
        }

        List<string[]> file_oshibki = new List<string[]>();
        public List<string[]> Get_Oshibki
        {
            get { return file_oshibki; }
        }

        protected internal List<string> Count_Oshibki_to_ListBox()
        {
            string[] files_full_paths = Directory.GetFiles(path_to_oshibki, "*.txt")
                                     .Select(Path.GetFullPath)
                                     .ToArray();

            file_names = new DirectoryInfo(path_to_oshibki).GetFiles("*.xlsx");

            List<string> file_c = new List<string>();
            

            //Начиная с 1,а не 0 т.к. нулевое - образец!!!
            for (int i = 0; i < files_full_paths.Count(); i++)
            {
                file_c.Add("Название файла: " + file_names[i]);

                file_oshibki.Add(File.ReadAllLines(files_full_paths[i]));
            }
            return file_c;
        }

        #region Переменные для работы с отчетами
        string text_connecta, data_plat_BD = "", vid_plat, UK = "", UK_BD, Fond_BD, chet_UK = "", fond, valuta, agent, naznach, nazv_vipiski, 
            nazv_otch, type_raspor, nomer_lic_UK, nomer_lic_fond;
        int nomer_PP;
        decimal summa, summa_deb = 0, summa_kred = 0, vxod_saldo_kred = 0, oborot_deb = 0, oborot_kred = 0, isxod_saldo_deb = 0, isxod_saldo_kred = 0;
        bool otmetka;
        int kontrol_kol_strok = 0, nomer_iteracii = 0;
        
        string newfile_name2 = "";

        //СЧА
        decimal pole_010 = 0, pole_011 = 0, pole_012 = 0, pole_160 = 0, pole_300 = 0, pole_310 = 0, pole_270 = 0,
            pole_330 = 0, pole_400 = 0, pole_600 = 0, pole_170 = 0;
        int kontrol_oplat_spec_dep = 2, kontrol_oplat_spec_reg = 2;
        double kol_pay = new double(), pole_500 = new double(), pay_UK_now = new double();
        #endregion

        #region Создание отчетов
        protected internal OleDbCommand Connect_to_BD_for_otchet(string connecta_text)     //ЭТО МЕТОД соединения с БД
        {
            OleDbCommand CommandBD = new OleDbCommand();                                      //команда, через которую все делается
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DataBaseMy.accdb";
            OleDbConnection conn = new OleDbConnection(connectionString);                       //новое подключение к БД
            CommandBD.Connection = conn;                                                      //соединение с бд
            conn.Open();
            CommandBD.CommandText = connecta_text;
            return CommandBD;
        }

        protected internal void InsertData_vyp(string data_v_boxe, int nomer_proxoda)          //ЭТО МЕТОД записи/обновления БД
        {
            string text_connecta = "SELECT * FROM Платежные_поручения";
            Connect_to_BD_for_otchet(text_connecta);
            var CommandBD = Connect_to_BD_for_otchet(text_connecta);
            var conn = CommandBD.Connection;

            #region //запись в таблицу выписка за день
            if (nomer_proxoda == 1)
            {
                CommandBD.CommandText = "INSERT INTO [Выписка_день] ([Название_выписки], [Дата_отчета], [Исходящее_cальдо_кредит], [Дата_отправки]) VALUES ('" + nazv_vipiski + "', '" + data_v_boxe + "', '" + isxod_saldo_kred + "', '" + data_v_boxe + "')";
                CommandBD.ExecuteNonQuery();
            }
            else if (nomer_proxoda > 1)
            {
                string Update_vipiska_day = "UPDATE Выписка_день SET [Исходящее_cальдо_кредит] = ? WHERE [Название_выписки] = ?";
                using (OleDbCommand CommandBDParams = new OleDbCommand(Update_vipiska_day, conn))
                {
                    CommandBDParams.Parameters.Add("@Q1", OleDbType.Decimal).Value = isxod_saldo_kred;
                    CommandBDParams.Parameters.Add("@Q2", OleDbType.Char).Value = nazv_vipiski;
                    CommandBDParams.ExecuteNonQuery();
                }
            }
            #endregion
            string Update_vipiska_day_every_time = "UPDATE Выписка_день SET [Отметка_об_отправке] = ? WHERE [Название_выписки] = ?";
            using (OleDbCommand CommandBDParams = new OleDbCommand(Update_vipiska_day_every_time, conn))
            {
                CommandBDParams.Parameters.Add("@Q1", OleDbType.Boolean).Value = true;
                CommandBDParams.Parameters.Add("@Q2", OleDbType.Char).Value = nazv_vipiski;
                CommandBDParams.ExecuteNonQuery();
            }

            #region //Запись всего в таблицу БД данных для выписки
            CommandBD.CommandText = "INSERT INTO [Данные_день_выписки] ([Номер_пп], [Название_выписки], [Дата], [УК], [Счет_УК], [Фонд], [Валюта], [Дебет], [Кредит], [Входящее_сальдо_кредит], [Итого_оборотов_дебет], [Итого_оборотов_кредит], [Исходящее_сальдо_дебет], [Исходящее_сальдо_кредит], [Назначение], [Контрагент]) VALUES ('" +
                nomer_PP + "', '" + nazv_vipiski + "', '" + data_v_boxe + "', '" + UK + "', '" + chet_UK + "', '" +
                fond + "', '" + valuta + "', '" + summa_deb + "', '" + summa_kred + "', '" + vxod_saldo_kred + "', '" + oborot_deb + "', '" + oborot_kred + "', '" +
                isxod_saldo_deb + "','" + isxod_saldo_kred + "','" + naznach + "','" + agent + "')";
            CommandBD.ExecuteNonQuery();

            //обновление "учтено/не учтено"
            string Update_PP = "UPDATE Платежные_поручения SET [Отметка_учета_вып] = ?, [Дата_учета] = ? WHERE [Номер_пп] = ?";
            using (OleDbCommand CommandBDParams = new OleDbCommand(Update_PP, conn))
            {
                CommandBDParams.Parameters.Add("@Q1", OleDbType.Boolean).Value = true;
                CommandBDParams.Parameters.Add("@Q2", OleDbType.Date).Value = DateTime.Now.ToShortDateString();
                CommandBDParams.Parameters.Add("@Q3", OleDbType.Integer).Value = nomer_PP;
                CommandBDParams.ExecuteNonQuery();
            }
            //изменение в таблице УК сальдо
            string Update_UK = "UPDATE Список_Управляющих_компаний SET [Текущее_сальдо_кредит] = ? WHERE [Полное_наименование_УК] = ?";
            using (OleDbCommand CommandBDParams = new OleDbCommand(Update_UK, conn))
            {
                CommandBDParams.Parameters.Add("@Q1", OleDbType.Currency).Value = isxod_saldo_kred;
                CommandBDParams.Parameters.Add("@Q2", OleDbType.Char).Value = UK_BD;
                CommandBDParams.ExecuteNonQuery();
            }
            #endregion
            conn.Close();
        }

        protected internal void New_stroka_excel(string file_name_open)            //Метод создания новой строки
        {
            Excel.Sheets excelsheets;
            Excel.Range excelcells;
            Excel.Application objWorkExcel = new Excel.Application();
            //Открываем уже новый (созданный файл, а не образец)
            Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(file_name_open, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelsheets = objWorkBook.Worksheets;                       //Получаем массив ссылок на листы выбранной книги
            Excel.Worksheet excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);  //Получаем ссылку на лист 1

            //создать новую строку !!! сверху !!! выбранной, затем копируем старую на новое место со старыми параметрами
            Excel.Range old_stroka = excelworksheet.get_Range("A10", "K10");
            old_stroka.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
            excelcells = excelworksheet.get_Range("A11", "K11");                            // Устанавливаем ссылку на ячейку A1
            excelcells.Copy(Type.Missing);
            excelcells = excelworksheet.get_Range("A10", "K10");                            // Устанавливаем ссылку на ячейку A1
            excelcells.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
            //Таким образом всегда запись идет в 10 строку

            objWorkExcel.DisplayAlerts = false;
            objWorkBook.SaveAs(@file_name_open, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            objWorkExcel.Quit();
        }

        protected internal void Zapis_v_excel_vipiska_day(string nazvanie)          //метод записи в excel
        {
            string table_name = "SELECT * FROM Данные_день_выписки";
            Connect_to_BD_for_otchet(table_name);
            var CommandBD = Connect_to_BD_for_otchet(table_name);
            var conn = CommandBD.Connection;
            OleDbDataReader dr1 = CommandBD.ExecuteReader();

            var obrazec_directory = new DirectoryInfo(@".\\Отчеты_на_отправку");
            string obrazec_name = "", newfile_n = "", newfile_name = "", newfile_name_arh = "";
            foreach (FileInfo file in obrazec_directory.GetFiles())
            {
                if (file.Name.Contains("Образец_выписка"))
                {
                    obrazec_name = file.FullName;
                    newfile_n = file.DirectoryName;
                }
            }

            Excel.Sheets excelsheets;
            Excel.Range excelcells;
            Excel.Application objWorkExcel = new Excel.Application();                   //подключим excel

            while (dr1.Read())
            {
                string name_vip_BD = dr1.GetString(1);
                bool otmetka_ucheta = dr1.GetBoolean(16);
                if ((name_vip_BD == nazvanie) && (otmetka_ucheta == false))
                {
                    //MessageBox.Show(Convert.ToString(kontrol_kol_strok), "kol_strok");    //проверка
                    #region                    Первая строка
                    if (kontrol_kol_strok == 0)
                    {
                        Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(obrazec_name, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        excelsheets = objWorkBook.Worksheets;                       //Получаем массив ссылок на листы выбранной книги
                        Excel.Worksheet excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);  //Получаем ссылку на лист 1

                        //Название выписки
                        excelcells = excelworksheet.get_Range("E1", Type.Missing);  //Выбираем ячейку для вывода
                        excelcells.Value2 = nazvanie;                                       //Записываем

                        //Наименование УК
                        excelcells = excelworksheet.get_Range("E2", Type.Missing);
                        excelcells.Value2 = UK;

                        //№ счета
                        excelcells = excelworksheet.get_Range("F3", Type.Missing);
                        excelcells.Value2 = chet_UK;

                        //Дата
                        excelcells = excelworksheet.get_Range("J3", Type.Missing);
                        excelcells.Value2 = data_plat_BD.Substring(0, 10);

                        //Наименование фонда
                        excelcells = excelworksheet.get_Range("E4", Type.Missing);
                        excelcells.Value2 = fond;

                        //Валюта
                        excelcells = excelworksheet.get_Range("C5", Type.Missing);
                        excelcells.Value2 = valuta;

                        //Входящее сальдо кредит
                        decimal vxod_saldo_kontrol_BD = dr1.GetDecimal(9);
                        Excel.Range Cell_vxod_saldo = excelworksheet.get_Range("F9");
                        string cell_vxod_s = Cell_vxod_saldo.Text.ToString();
                        decimal vxod_saldo_kontrol_Excel = Convert.ToDecimal(cell_vxod_s);
                        if (vxod_saldo_kontrol_BD > vxod_saldo_kontrol_Excel)
                        {
                            excelcells = excelworksheet.get_Range("F9", Type.Missing);
                            excelcells.Value2 = vxod_saldo_kontrol_BD;
                        }
                        else
                        {
                            //ничего не запишется
                        }

                        /*MessageBox.Show(Convert.ToString(nomer_PP), "nomer_PP");
                        MessageBox.Show(Convert.ToString(summa), "summa");
                        MessageBox.Show(Convert.ToString(summa_deb), "summa_d");
                        MessageBox.Show(Convert.ToString(summa_kred), "summa_k");*/

                        //поля п/п
                        //дата
                        excelcells = excelworksheet.get_Range("A10", Type.Missing);
                        excelcells.Value2 = data_plat_BD.Substring(0, 10);

                        //док. (№ п/п)
                        excelcells = excelworksheet.get_Range("C10", Type.Missing);
                        excelcells.Value2 = nomer_PP;

                        //Счет
                        excelcells = excelworksheet.get_Range("G10", Type.Missing);
                        excelcells.Value2 = chet_UK;

                        //Дебет
                        excelcells = excelworksheet.get_Range("H10", Type.Missing);
                        excelcells.Value2 = summa_deb;

                        //Кредит
                        excelcells = excelworksheet.get_Range("I10", Type.Missing);
                        excelcells.Value2 = summa_kred;

                        //Назначение
                        excelcells = excelworksheet.get_Range("J10", Type.Missing);
                        excelcells.Value2 = naznach;

                        //Контрагент
                        excelcells = excelworksheet.get_Range("K10", Type.Missing);
                        excelcells.Value2 = agent;

                        //итого оборотов дебет
                        decimal itogo_oborot_deb_kontrol_BD = dr1.GetDecimal(10);
                        int cell_number_oborot_deb = 11 + kontrol_kol_strok;                                    //изменяемая строка ячейки
                        string cell_name_oborot_deb = "E" + cell_number_oborot_deb;                             //изменяемое название ячейки
                        Excel.Range Cell_itogo_oborot_deb = excelworksheet.get_Range(cell_name_oborot_deb);
                        string cell_it_ob_deb = Cell_itogo_oborot_deb.Text.ToString();
                        decimal itogo_oborot_deb_kontrol_Excel = Convert.ToDecimal(cell_it_ob_deb);
                        if (itogo_oborot_deb_kontrol_BD > itogo_oborot_deb_kontrol_Excel)
                        {
                            excelcells = excelworksheet.get_Range(cell_name_oborot_deb, Type.Missing);
                            excelcells.Value2 = itogo_oborot_deb_kontrol_BD;
                        }
                        else
                        {
                            //nothing
                        }

                        //итого оборотов кредит
                        decimal itogo_oborot_kred_kontrol_BD = dr1.GetDecimal(11);
                        int cell_number_oborot_kred = 11 + kontrol_kol_strok;                                    //изменяемая строка ячейки
                        string cell_name_oborot_kred = "H" + cell_number_oborot_kred;                             //изменяемое название ячейки
                        Excel.Range Cell_itogo_oborot_kred = excelworksheet.get_Range(cell_name_oborot_kred);
                        string cell_it_ob_kred = Cell_itogo_oborot_kred.Text.ToString();
                        decimal itogo_oborot_kred_kontrol_Excel = Convert.ToDecimal(cell_it_ob_kred);
                        if (itogo_oborot_kred_kontrol_BD > itogo_oborot_kred_kontrol_Excel)
                        {
                            excelcells = excelworksheet.get_Range(cell_name_oborot_kred, Type.Missing);
                            excelcells.Value2 = itogo_oborot_kred_kontrol_BD;
                        }
                        else
                        {
                            //nothing
                        }

                        //исходящее сальдо дебет
                        decimal isx_saldo_deb_kontrol_BD = dr1.GetDecimal(12);
                        int cell_number_isx_saldo_deb = 12 + kontrol_kol_strok;                                    //изменяемая строка ячейки
                        string cell_name_isx_saldo_deb = "E" + cell_number_isx_saldo_deb;                             //изменяемое название ячейки
                        Excel.Range Cell_isx_saldo_deb = excelworksheet.get_Range(cell_name_isx_saldo_deb);
                        string cell_isx_s__deb = Cell_isx_saldo_deb.Text.ToString();
                        decimal isx_saldo_deb_kontrol_Excel = Convert.ToDecimal(cell_isx_s__deb);
                        if (isx_saldo_deb_kontrol_BD > isx_saldo_deb_kontrol_Excel)
                        {
                            excelcells = excelworksheet.get_Range(cell_name_isx_saldo_deb, Type.Missing);
                            excelcells.Value2 = isx_saldo_deb_kontrol_BD;
                        }
                        else
                        {
                            //nothing
                        }

                        //исходящее сальдо кредит
                        decimal isx_saldo_kred_kontrol_BD = dr1.GetDecimal(13);
                        int cell_number_isx_saldo_kred = 12 + kontrol_kol_strok;                                    //изменяемая строка ячейки
                        string cell_name_isx_saldo_kred = "H" + cell_number_isx_saldo_kred;                             //изменяемое название ячейки
                        Excel.Range Cell_isx_saldo_kred = excelworksheet.get_Range(cell_name_isx_saldo_kred);
                        string cell_isx_s__kred = Cell_isx_saldo_kred.Text.ToString();
                        decimal isx_saldo_kred_kontrol_Excel = Convert.ToDecimal(cell_isx_s__kred);
                        if (isx_saldo_kred_kontrol_BD < isx_saldo_kred_kontrol_Excel)
                        {
                            excelcells = excelworksheet.get_Range(cell_name_isx_saldo_kred, Type.Missing);
                            excelcells.Value2 = isx_saldo_kred_kontrol_BD;
                        }
                        else
                        {
                            //nothing
                        }

                        //Сохранение файла
                        newfile_name = newfile_n + "\\" + nazv_vipiski + ".xls";
                        newfile_name_arh = path_to_arh_otch + "\\" + nazv_vipiski + ".xls";
                        //MessageBox.Show(newfile_name, "newfile_name");

                        objWorkExcel.DisplayAlerts = false;
                        objWorkBook.SaveAs(@newfile_name, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        //objWorkBook.Close(Type.Missing, obrazec_name, Type.Missing);
                        objWorkExcel.Quit();
                        File.Copy(newfile_name, newfile_name_arh, true);

                        //kontrol_kol_strok++;
                        //MessageBox.Show("1 строка готова");

                        //постановка отметки об учете
                        string Update_PP = "UPDATE Данные_день_выписки SET [Отметка_об_учете] = ? WHERE [Название_выписки] = ? AND [Номер_пп] = ?";
                        using (OleDbCommand CommandBDParams = new OleDbCommand(Update_PP, conn))
                        {
                            CommandBDParams.Parameters.Add("@Q1", OleDbType.Boolean).Value = true;
                            CommandBDParams.Parameters.Add("@Q2", OleDbType.Char).Value = nazv_vipiski;
                            CommandBDParams.Parameters.Add("@Q3", OleDbType.Integer).Value = nomer_PP;
                            CommandBDParams.ExecuteNonQuery();
                        }
                    }
                    #endregion
                    #region ВТОРАЯ И ПОСЛЕДУЮЩИЕ СТРОКИ!!!
                    if (kontrol_kol_strok > 0)
                    {
                        /*//создать новую строку !!! сверху !!! выбранной, затем копируем старую на новое место со старыми параметрами
                        Excel.Range old_stroka = excelworksheet.get_Range("A10", "K10");
                        old_stroka.EntireRow.Insert(Excel.XlInsertShiftDirection.xlShiftDown);
                        excelcells = excelworksheet.get_Range("A11", "K11");                            // Устанавливаем ссылку на ячейку A1
                        excelcells.Copy(Type.Missing);
                        excelcells = excelworksheet.get_Range("A10", "K10");                            // Устанавливаем ссылку на ячейку A1
                        excelcells.PasteSpecial(Excel.XlPasteType.xlPasteAll, Excel.XlPasteSpecialOperation.xlPasteSpecialOperationNone, false, false);
                        //Таким образом всегда запись идет в 10 строку*/
                        newfile_name2 = newfile_n + "\\" + nazv_vipiski + ".xls";
                        New_stroka_excel(newfile_name2);     //используем метод создания новой строки

                        //Открываем уже новый (созданный файл, а не образец)
                        Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(newfile_name2, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        excelsheets = objWorkBook.Worksheets;                       //Получаем массив ссылок на листы выбранной книги
                        Excel.Worksheet excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);  //Получаем ссылку на лист 1

                        //Входящее сальдо кредит
                        decimal vxod_saldo_kontrol_BD = dr1.GetDecimal(9);
                        Excel.Range Cell_vxod_saldo = excelworksheet.get_Range("F9");
                        string cell_vxod_s = Cell_vxod_saldo.Text.ToString();
                        decimal vxod_saldo_kontrol_Excel = Convert.ToDecimal(cell_vxod_s);
                        if (vxod_saldo_kontrol_BD > vxod_saldo_kontrol_Excel)
                        {
                            excelcells = excelworksheet.get_Range("F9", Type.Missing);
                            excelcells.Value2 = vxod_saldo_kontrol_BD;
                        }
                        else
                        {
                            //ничего не запишется
                        }

                        //поля п/п
                        //дата
                        //string d = "A" + cell_numeric;
                        excelcells = excelworksheet.get_Range("A10", Type.Missing);
                        excelcells.Value2 = data_plat_BD.Substring(0, 10);

                        //док. (№ п/п)
                        //string dok = "C" + cell_numeric;
                        excelcells = excelworksheet.get_Range("C10", Type.Missing);
                        excelcells.Value2 = nomer_PP;

                        //Счет
                        //string c = "G" + cell_numeric;
                        excelcells = excelworksheet.get_Range("G10", Type.Missing);
                        excelcells.Value2 = chet_UK;

                        //Дебет
                        //string sd = "H" + cell_numeric;
                        excelcells = excelworksheet.get_Range("H10", Type.Missing);
                        excelcells.Value2 = summa_deb;

                        //Кредит
                        //string sk = "I" + cell_numeric;
                        excelcells = excelworksheet.get_Range("I10", Type.Missing);
                        excelcells.Value2 = summa_kred;

                        //Назначение
                        //string n = "J" + cell_numeric;
                        excelcells = excelworksheet.get_Range("J10", Type.Missing);
                        excelcells.Value2 = naznach;

                        //Контрагент
                        //string a = "K" + cell_numeric;
                        excelcells = excelworksheet.get_Range("K10", Type.Missing);
                        excelcells.Value2 = agent;

                        //итого оборотов дебет
                        decimal itogo_oborot_deb_kontrol_BD = dr1.GetDecimal(10);
                        int cell_number_oborot_deb = 11 + kontrol_kol_strok;                                    //изменяемая строка ячейки
                        string cell_name_oborot_deb = "E" + cell_number_oborot_deb;                             //изменяемое название ячейки
                        Excel.Range Cell_itogo_oborot_deb = excelworksheet.get_Range(cell_name_oborot_deb);
                        string cell_it_ob_deb = Cell_itogo_oborot_deb.Text.ToString();
                        decimal itogo_oborot_deb_kontrol_Excel = Convert.ToDecimal(cell_it_ob_deb);
                        if (itogo_oborot_deb_kontrol_BD > itogo_oborot_deb_kontrol_Excel)
                        {
                            excelcells = excelworksheet.get_Range(cell_name_oborot_deb, Type.Missing);
                            excelcells.Value2 = itogo_oborot_deb_kontrol_BD;
                        }
                        else
                        {
                            //nothing
                        }

                        //итого оборотов кредит
                        decimal itogo_oborot_kred_kontrol_BD = dr1.GetDecimal(11);
                        int cell_number_oborot_kred = 11 + kontrol_kol_strok;                                    //изменяемая строка ячейки
                        string cell_name_oborot_kred = "H" + cell_number_oborot_kred;                             //изменяемое название ячейки
                        Excel.Range Cell_itogo_oborot_kred = excelworksheet.get_Range(cell_name_oborot_kred);
                        string cell_it_ob_kred = Cell_itogo_oborot_kred.Text.ToString();
                        decimal itogo_oborot_kred_kontrol_Excel = Convert.ToDecimal(cell_it_ob_kred);
                        if (itogo_oborot_kred_kontrol_BD > itogo_oborot_kred_kontrol_Excel)
                        {
                            excelcells = excelworksheet.get_Range(cell_name_oborot_kred, Type.Missing);
                            excelcells.Value2 = itogo_oborot_kred_kontrol_BD;
                        }
                        else
                        {
                            //nothing
                        }

                        //исходящее сальдо дебет
                        decimal isx_saldo_deb_kontrol_BD = dr1.GetDecimal(12);
                        int cell_number_isx_saldo_deb = 12 + kontrol_kol_strok;                                    //изменяемая строка ячейки
                        string cell_name_isx_saldo_deb = "E" + cell_number_isx_saldo_deb;                             //изменяемое название ячейки
                        Excel.Range Cell_isx_saldo_deb = excelworksheet.get_Range(cell_name_isx_saldo_deb);
                        string cell_isx_s__deb = Cell_isx_saldo_deb.Text.ToString();
                        decimal isx_saldo_deb_kontrol_Excel = Convert.ToDecimal(cell_isx_s__deb);
                        if (isx_saldo_deb_kontrol_BD >= isx_saldo_deb_kontrol_Excel)
                        {
                            excelcells = excelworksheet.get_Range(cell_name_isx_saldo_deb, Type.Missing);
                            excelcells.Value2 = isx_saldo_deb_kontrol_BD;
                        }
                        else
                        {
                            //nothing
                        }

                        //исходящее сальдо кредит
                        decimal isx_saldo_kred_kontrol_BD = dr1.GetDecimal(13);
                        int cell_number_isx_saldo_kred = 12 + kontrol_kol_strok;                                    //изменяемая строка ячейки
                        string cell_name_isx_saldo_kred = "H" + cell_number_isx_saldo_kred;                             //изменяемое название ячейки
                        Excel.Range Cell_isx_saldo_kred = excelworksheet.get_Range(cell_name_isx_saldo_kred);
                        string cell_isx_s__kred = Cell_isx_saldo_kred.Text.ToString();
                        decimal isx_saldo_kred_kontrol_Excel = Convert.ToDecimal(cell_isx_s__kred);
                        if (isx_saldo_kred_kontrol_BD <= isx_saldo_kred_kontrol_Excel)
                        {
                            excelcells = excelworksheet.get_Range(cell_name_isx_saldo_kred, Type.Missing);
                            excelcells.Value2 = isx_saldo_kred_kontrol_BD;
                        }
                        else
                        {
                            //nothing
                        }

                        //Пересохранение
                        objWorkExcel.DisplayAlerts = false;
                        objWorkBook.SaveAs(@newfile_name2, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        objWorkExcel.Quit();
                        newfile_name_arh = path_to_arh_otch + "\\" + nazv_vipiski + ".xls";
                        File.Copy(newfile_name2, newfile_name_arh, true);

                        //kontrol_kol_strok++;

                        //постановка отметки об учете
                        string Update_PP = "UPDATE Данные_день_выписки SET [Отметка_об_учете] = ? WHERE [Название_выписки] = ? AND [Номер_пп] = ?";
                        using (OleDbCommand CommandBDParams = new OleDbCommand(Update_PP, conn))
                        {
                            CommandBDParams.Parameters.Add("@Q1", OleDbType.Boolean).Value = true;
                            CommandBDParams.Parameters.Add("@Q2", OleDbType.Char).Value = nazv_vipiski;
                            CommandBDParams.Parameters.Add("@Q3", OleDbType.Integer).Value = nomer_PP;
                            CommandBDParams.ExecuteNonQuery();
                        }
                    }
                    #endregion

                    kontrol_kol_strok++;
                }
                else
                {
                    objWorkExcel.Quit();
                }
                //kontrol_kol_strok++;
            }
        }

        protected internal void InsertData_Scha_day(string data_v_boxe, int nomer_proxoda)
        {
            string text_connecta = "SELECT * FROM Платежные_поручения";
            Connect_to_BD_for_otchet(text_connecta);
            var CommandBD = Connect_to_BD_for_otchet(text_connecta);
            var conn = CommandBD.Connection;

            #region //запись в таблицу сча за день
            if (nomer_proxoda == 0)
            {
                CommandBD.CommandText = "INSERT INTO [СЧА_день] ([Название_отчета], [Дата], [Дата_отчета], [СЧА], [Кол_паев], [Цена_пая], [УК], [Фонд], [Дата_отправки], [ДС_всего], [ДС_руб], [ДС_ин], [Актив_всего], [Обяз_всего], [Обяз_кред_задолжн], [Обяз_резерв], [Недвиж_РФ], [Недвиж_не_РФ]) VALUES ('" +
                    nazv_otch + "', '" + DateTime.Now.ToShortDateString() + "', '" + data_v_boxe + "', '" + pole_400 + "', '" + pole_500 + "', '" +
                    pole_600 + "', '" + UK + "','" + fond + "', '" + DateTime.Now.ToShortDateString() + "', '" + pole_010 + "', '" + pole_011 + "', '" +
                    pole_012 + "', '" + pole_270 + "', '" + pole_330 + "', '" + pole_300 + "', '" + pole_310 + "', '" + pole_160 + "', '" + pole_170 + "')";
                CommandBD.ExecuteNonQuery();
            }
            else if (nomer_proxoda > 0)
            {
                string Update_vipiska_day = "UPDATE СЧА_день SET [СЧА] = ?, [Кол_паев] = ?, [Цена_пая] = ?, [ДС_всего] = ?, [ДС_руб] = ?, [ДС_ин] = ?, [Актив_всего] = ?, [Обяз_всего] = ?, [Обяз_кред_задолжн] = ?, [Обяз_резерв] = ?, [Недвиж_РФ] = ?, [Недвиж_не_РФ] = ? WHERE [Название_отчета] = ?";
                using (OleDbCommand CommandBDParams = new OleDbCommand(Update_vipiska_day, conn))
                {
                    CommandBDParams.Parameters.Add("@Q1", OleDbType.Decimal).Value = pole_400;
                    CommandBDParams.Parameters.Add("@Q2", OleDbType.Double).Value = pole_500;
                    CommandBDParams.Parameters.Add("@Q3", OleDbType.Double).Value = pole_600;
                    CommandBDParams.Parameters.Add("@Q4", OleDbType.Decimal).Value = pole_010;
                    CommandBDParams.Parameters.Add("@Q5", OleDbType.Decimal).Value = pole_011;
                    CommandBDParams.Parameters.Add("@Q6", OleDbType.Decimal).Value = pole_012;
                    CommandBDParams.Parameters.Add("@Q7", OleDbType.Decimal).Value = pole_270;
                    CommandBDParams.Parameters.Add("@Q8", OleDbType.Decimal).Value = pole_330;
                    CommandBDParams.Parameters.Add("@Q9", OleDbType.Decimal).Value = pole_300;
                    CommandBDParams.Parameters.Add("@Q10", OleDbType.Decimal).Value = pole_310;
                    CommandBDParams.Parameters.Add("@Q11", OleDbType.Decimal).Value = pole_160;
                    CommandBDParams.Parameters.Add("@Q12", OleDbType.Decimal).Value = pole_170;
                    CommandBDParams.Parameters.Add("@Q3", OleDbType.Char).Value = nazv_otch;
                    CommandBDParams.ExecuteNonQuery();
                }
            }
            #endregion
            string Update_scha_day_every_time = "UPDATE СЧА_день SET [Отметка_об_отправке] = ? WHERE [Название_отчета] = ?";
            using (OleDbCommand CommandBDParams = new OleDbCommand(Update_scha_day_every_time, conn))
            {
                CommandBDParams.Parameters.Add("@Q1", OleDbType.Boolean).Value = true;
                CommandBDParams.Parameters.Add("@Q2", OleDbType.Char).Value = nazv_otch;
                CommandBDParams.ExecuteNonQuery();
            }
            #region //запись в таблицу данные для сча день
            CommandBD.CommandText = "INSERT INTO [Данные_день_СЧА] ([Номер_пп], [Название_отчета], [Дата], [УК], [Фонд], [Активы], [Обязательства], [СЧА], [Кол_паев], [Цена_пая], [Назначение], [Контрагент]) VALUES ('" +
                nomer_PP + "', '" + nazv_otch + "', '" + Convert.ToDateTime(data_plat_BD).ToShortDateString() + "', '" + UK + "', '" + fond +
                "', '" + pole_270 + "', '" + pole_330 + "', '" + pole_400 + "', '" + kol_pay + "', '"+ pole_600 +"', '" + 
                naznach + "', '" + agent + "')";
            CommandBD.ExecuteNonQuery();

            //обновление "учтено/не учтено"
            string Update_PP = "UPDATE Платежные_поручения SET [Отметка_учета_сча] = ?, [Дата_учета] = ? WHERE [Номер_пп] = ?";
            using (OleDbCommand CommandBDParams = new OleDbCommand(Update_PP, conn))
            {
                CommandBDParams.Parameters.Add("@Q1", OleDbType.Boolean).Value = true;
                CommandBDParams.Parameters.Add("@Q2", OleDbType.Date).Value = DateTime.Now.ToShortDateString();
                CommandBDParams.Parameters.Add("@Q3", OleDbType.Integer).Value = nomer_PP;
                CommandBDParams.ExecuteNonQuery();
            }
            //изменение в таблице УК сальдо
            string Update_UK = "UPDATE Список_Управляющих_компаний SET [Текущее_кол_паев] = ? WHERE [Полное_наименование_УК] = ?";
            using (OleDbCommand CommandBDParams = new OleDbCommand(Update_UK, conn))
            {
                CommandBDParams.Parameters.Add("@Q1", OleDbType.Double).Value = pole_500;
                CommandBDParams.Parameters.Add("@Q2", OleDbType.Char).Value = UK_BD;
                CommandBDParams.ExecuteNonQuery();
            }
            #endregion
        }

        protected internal void Zapis_v_excel_Scha_day_today(string nazvanie, string vid_plat_pp, decimal pole_010,decimal pole_011,
            decimal pole_012,decimal pole_160,decimal pole_300,decimal pole_310,decimal pole_270,
           decimal pole_330,decimal pole_400,decimal pole_600,decimal pole_170)
        {
            string table_name = "SELECT * FROM Данные_день_СЧА";
            Connect_to_BD_for_otchet(table_name);
            var CommandBD = Connect_to_BD_for_otchet(table_name);
            var conn = CommandBD.Connection;
            OleDbDataReader dr1 = CommandBD.ExecuteReader();

            var obrazec_directory = new DirectoryInfo(@".\\Отчеты_на_отправку");
            string obrazec_name = "", newfile_n = "", newfile_name = "", newfile_name_arh = "";
            foreach (FileInfo file in obrazec_directory.GetFiles())
            {
                if (file.Name.Contains("Образец_сча_день"))
                {
                    obrazec_name = file.FullName;
                    newfile_n = file.DirectoryName;
                }
            }
            
            Excel.Sheets excelsheets;
            Excel.Range excelcells;
            Excel.Application objWorkExcel = new Excel.Application();                   //подключим excel

            while (dr1.Read())
            {
                string name_otch_BD = dr1.GetString(1);
                DateTime date_pole = dr1.GetDateTime(2);
                bool otmetka_ucheta = dr1.GetBoolean(12);

                #region //первый проход
                if (kontrol_kol_strok == 0)
                {
                    if ((name_otch_BD == nazvanie) && (otmetka_ucheta == false))
                    {
                        newfile_name2 = newfile_n + "\\" + nazv_otch + ".xls";
                        Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(newfile_name2, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        excelsheets = objWorkBook.Worksheets;                       //Получаем массив ссылок на листы выбранной книги
                        Excel.Worksheet excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);  //Получаем ссылку на лист 1

                        #region //запись текущего дня
                        {
                            #region //Название отчета
                            excelcells = excelworksheet.get_Range("D1", Type.Missing);  //Выбираем ячейку для вывода
                            excelcells.Value2 = nazvanie;                                       //Записываем
                            #endregion
                            #region //Наименование УК
                            excelcells = excelworksheet.get_Range("H9", Type.Missing);
                            excelcells.Value2 = UK;
                            #endregion
                            #region //Лицензия УК
                            excelcells = excelworksheet.get_Range("K9", Type.Missing);
                            excelcells.Value2 = nomer_lic_UK;
                            #endregion
                            #region //Наименование фонда
                            excelcells = excelworksheet.get_Range("A9", Type.Missing);
                            excelcells.Value2 = fond;
                            #endregion
                            #region //лицензия фонда
                            excelcells = excelworksheet.get_Range("D9", Type.Missing);
                            excelcells.Value2 = nomer_lic_fond;
                            #endregion
                            #region //Дата текущая
                            excelcells = excelworksheet.get_Range("I14", Type.Missing);
                            excelcells.Value2 = data_plat_BD.Substring(0, 10);
                            #endregion

                            #region //поля отчета
                            decimal a = new decimal(), b = new decimal(), c = new decimal(), d = new decimal();
                            #region //ДС всего, рубли и валюта
                            Excel.Range Cell_DC_all = excelworksheet.get_Range("L16");
                            string cell_dc_all = Cell_DC_all.Text.ToString();
                            decimal DC_all_yes = Convert.ToDecimal(cell_dc_all);
                            excelcells = excelworksheet.get_Range("I16", Type.Missing);
                            if (vid_plat_pp == "Дебет")    //платят нам
                            {
                                excelcells.Value2 = DC_all_yes + pole_010;
                            }
                            if (vid_plat_pp == "Кредит")   //платим мы
                            {
                                excelcells.Value2 = DC_all_yes - pole_010;
                            }
                            Excel.Range Cell_DC_rub = excelworksheet.get_Range("L18");
                            string cell_dc_rub = Cell_DC_rub.Text.ToString();
                            decimal DC_rub_yes = Convert.ToDecimal(cell_dc_rub);
                            excelcells = excelworksheet.get_Range("I18", Type.Missing);
                            if (vid_plat_pp == "Дебет")    //платят нам
                            {
                                excelcells.Value2 = DC_rub_yes + pole_011;
                            }
                            if (vid_plat_pp == "Кредит")   //платим мы
                            {
                                excelcells.Value2 = DC_rub_yes - pole_011;
                            }
                            Excel.Range Cell_DC_in = excelworksheet.get_Range("L19");
                            string cell_dc_in = Cell_DC_in.Text.ToString();
                            decimal DC_in_yes = Convert.ToDecimal(cell_dc_in);
                            excelcells = excelworksheet.get_Range("I19", Type.Missing);
                            if (vid_plat_pp == "Дебет")    //платят нам
                            {
                                excelcells.Value2 = DC_in_yes + pole_012;
                            }
                            if (vid_plat_pp == "Кредит")   //платим мы
                            {
                                excelcells.Value2 = DC_in_yes - pole_012;
                            }
                            #endregion
                            #region //Недвиж. в РФ + вне РФ
                            Excel.Range Cell_nedv_rus = excelworksheet.get_Range("L53");
                            string cell_nedv_rus = Cell_nedv_rus.Text.ToString();
                            decimal Nedv_rus_yes = Convert.ToDecimal(cell_nedv_rus);
                            excelcells = excelworksheet.get_Range("I53", Type.Missing);
                            if (vid_plat_pp == "Дебет")    //платят нам
                            {
                                excelcells.Value2 = Nedv_rus_yes + pole_160;
                            }
                            if (vid_plat_pp == "Кредит")   //платим мы
                            {
                                excelcells.Value2 = Nedv_rus_yes - pole_160;
                            }
                            Excel.Range Cell_nedv_in = excelworksheet.get_Range("L57");
                            string cell_nedv_in = Cell_nedv_in.Text.ToString();
                            decimal Nedv_in_yes = Convert.ToDecimal(cell_nedv_in);
                            excelcells = excelworksheet.get_Range("I57", Type.Missing);
                            if (vid_plat_pp == "Дебет")    //платят нам
                            {
                                excelcells.Value2 = Nedv_in_yes + pole_170;
                            }
                            if (vid_plat_pp == "Кредит")   //платим мы
                            {
                                excelcells.Value2 = Nedv_in_yes - pole_170;
                            }
                            #endregion
                            #region //Сумма активов
                            excelcells = excelworksheet.get_Range("I96", Type.Missing);
                            if (vid_plat_pp == "Дебет")
                            {
                                excelcells.Value2 = (DC_all_yes + pole_010) + (Nedv_rus_yes + pole_160) + (Nedv_in_yes + pole_170);
                                a = (DC_all_yes + pole_010) + (Nedv_rus_yes + pole_160) + (Nedv_in_yes + pole_170);
                            }
                            if (vid_plat_pp == "Кредит")
                            {
                                excelcells.Value2 = (DC_all_yes - pole_010) + (Nedv_rus_yes - pole_160) + (Nedv_in_yes - pole_170);
                                b = (DC_all_yes - pole_010) + (Nedv_rus_yes - pole_160) + (Nedv_in_yes - pole_170);
                            }
                            #endregion
                            #region //Кред. задолженность
                            Excel.Range Cell_kred_z = excelworksheet.get_Range("L100");
                            string cell_kred_z = Cell_kred_z.Text.ToString();
                            decimal kred_z = Convert.ToDecimal(cell_kred_z);
                            excelcells = excelworksheet.get_Range("I100", Type.Missing);
                            excelcells.Value2 = kred_z + pole_300;
                            #endregion
                            #region //Резерв на вознагр.
                            Excel.Range Cell_rez = excelworksheet.get_Range("L101");
                            string cell_rez = Cell_rez.Text.ToString();
                            decimal rezerv = Convert.ToDecimal(cell_rez);
                            excelcells = excelworksheet.get_Range("I101", Type.Missing);
                            excelcells.Value2 = rezerv + pole_310;
                            #endregion
                            #region //Сумма обязательств
                            excelcells = excelworksheet.get_Range("I105", Type.Missing);
                            excelcells.Value2 = pole_330;
                            #endregion
                            #region //СЧА
                            excelcells = excelworksheet.get_Range("I106", Type.Missing);
                            if (vid_plat_pp == "Дебет")
                            {
                                excelcells.Value2 = a - pole_330;
                                c = a - pole_330;
                            }
                            if (vid_plat_pp == "Кредит")
                            {
                                excelcells.Value2 = b - pole_330;
                                d = b - pole_330;
                            }
                            #endregion
                            #region //Кол. паев
                            excelcells = excelworksheet.get_Range("I107", Type.Missing);
                            excelcells.Value2 = pole_500;
                            #endregion
                            #region //Цена пая
                            excelcells = excelworksheet.get_Range("I109", Type.Missing);
                            if (vid_plat_pp == "Дебет")
                            {
                                excelcells.Value2 = c / Convert.ToDecimal(pole_500);
                            }
                            if (vid_plat_pp == "Кредит")
                            {
                                excelcells.Value2 = d / Convert.ToDecimal(pole_500);
                            }
                            #endregion

                            #endregion
                            
                            #region //Сохранение файла
                            newfile_name = newfile_n + "\\" + nazv_otch + ".xls";
                            newfile_name_arh = path_to_arh_otch + "\\" + nazv_otch + ".xls";
                            //MessageBox.Show(newfile_name, "newfile_name");

                            objWorkExcel.DisplayAlerts = false;
                            objWorkBook.SaveAs(@newfile_name, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            objWorkExcel.Quit();
                            File.Copy(newfile_name, newfile_name_arh, true);
                            #endregion

                            #region //постановка отметки об учете
                            string Update_PP = "UPDATE Данные_день_СЧА SET [Отметка_об_учете] = ? WHERE [Название_отчета] = ? AND [Номер_пп] = ?";
                            using (OleDbCommand CommandBDParams = new OleDbCommand(Update_PP, conn))
                            {
                                CommandBDParams.Parameters.Add("@Q1", OleDbType.Boolean).Value = true;
                                CommandBDParams.Parameters.Add("@Q2", OleDbType.Char).Value = nazv_otch;
                                CommandBDParams.Parameters.Add("@Q3", OleDbType.Integer).Value = nomer_PP;
                                CommandBDParams.ExecuteNonQuery();
                            }
                            #endregion
                        }
                        #endregion
                        kontrol_kol_strok++;
                    }
                    else
                    {
                        objWorkExcel.Quit();
                    }
                    
                }
                #endregion
                #region //следующий проход
                if (kontrol_kol_strok > 0)
                {
                    if ((name_otch_BD == nazvanie) && (otmetka_ucheta == false))
                    {
                        #region //открыть уже созданный отчет
                        newfile_name2 = newfile_n + "\\" + nazv_otch + ".xls";
                        Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(newfile_name2, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                        excelsheets = objWorkBook.Worksheets;                       //Получаем массив ссылок на листы выбранной книги
                        Excel.Worksheet excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);  //Получаем ссылку на лист 1
                        #endregion

                        #region //запись текущего дня
                        {
                            #region //Название отчета
                            excelcells = excelworksheet.get_Range("D1", Type.Missing);  //Выбираем ячейку для вывода
                            excelcells.Value2 = nazvanie;                                       //Записываем
                            #endregion
                            #region //Наименование УК
                            excelcells = excelworksheet.get_Range("H9", Type.Missing);
                            excelcells.Value2 = UK;
                            #endregion
                            #region //Лицензия УК
                            excelcells = excelworksheet.get_Range("K9", Type.Missing);
                            excelcells.Value2 = nomer_lic_UK;
                            #endregion
                            #region //Наименование фонда
                            excelcells = excelworksheet.get_Range("A9", Type.Missing);
                            excelcells.Value2 = fond;
                            #endregion
                            #region //лицензия фонда
                            excelcells = excelworksheet.get_Range("D9", Type.Missing);
                            excelcells.Value2 = nomer_lic_fond;
                            #endregion
                            #region //Дата текущая
                            excelcells = excelworksheet.get_Range("I14", Type.Missing);
                            excelcells.Value2 = data_plat_BD.Substring(0, 10);
                            #endregion

                            #region //поля отчета
                            #region //ДС всего, рубли и валюта
                            Excel.Range Cell_DC_all = excelworksheet.get_Range("L16");
                            string cell_dc_all = Cell_DC_all.Text.ToString();
                            decimal DC_all_yes = Convert.ToDecimal(cell_dc_all);
                            excelcells = excelworksheet.get_Range("I16", Type.Missing);
                            if (vid_plat_pp == "Дебет")    //платят нам
                            {
                                excelcells.Value2 = DC_all_yes + pole_010;
                                pole_010 = DC_all_yes + pole_010;
                            }
                            if (vid_plat_pp == "Кредит")   //платим мы
                            {
                                excelcells.Value2 = DC_all_yes - pole_010;
                                pole_010 = DC_all_yes - pole_010;
                            }
                            Excel.Range Cell_DC_rub = excelworksheet.get_Range("L18");
                            string cell_dc_rub = Cell_DC_rub.Text.ToString();
                            decimal DC_rub_yes = Convert.ToDecimal(cell_dc_rub);
                            excelcells = excelworksheet.get_Range("I18", Type.Missing);
                            if (vid_plat_pp == "Дебет")    //платят нам
                            {
                                excelcells.Value2 = DC_rub_yes + pole_011;
                                pole_011 = DC_rub_yes + pole_011;
                            }
                            if (vid_plat_pp == "Кредит")   //платим мы
                            {
                                excelcells.Value2 = DC_rub_yes - pole_011;
                                pole_011 = DC_rub_yes - pole_011;
                            }
                            Excel.Range Cell_DC_in = excelworksheet.get_Range("L19");
                            string cell_dc_in = Cell_DC_in.Text.ToString();
                            decimal DC_in_yes = Convert.ToDecimal(cell_dc_in);
                            excelcells = excelworksheet.get_Range("I19", Type.Missing);
                            if (vid_plat_pp == "Дебет")    //платят нам
                            {
                                excelcells.Value2 = DC_in_yes + pole_012;
                                pole_012 = DC_in_yes + pole_012;
                            }
                            if (vid_plat_pp == "Кредит")   //платим мы
                            {
                                excelcells.Value2 = DC_in_yes - pole_012;
                                pole_012 = DC_in_yes - pole_012;
                            }
                            #endregion
                            #region //Недвиж. в РФ + вне РФ
                            Excel.Range Cell_nedv_rus = excelworksheet.get_Range("L53");
                            string cell_nedv_rus = Cell_nedv_rus.Text.ToString();
                            decimal Nedv_rus_yes = Convert.ToDecimal(cell_nedv_rus);
                            excelcells = excelworksheet.get_Range("I53", Type.Missing);
                            if (vid_plat_pp == "Дебет")    //платят нам
                            {
                                excelcells.Value2 = Nedv_rus_yes + pole_160;
                                pole_160 = Nedv_rus_yes + pole_160;
                            }
                            if (vid_plat_pp == "Кредит")   //платим мы
                            {
                                excelcells.Value2 = Nedv_rus_yes - pole_160;
                                pole_160 = Nedv_rus_yes - pole_160;
                            }
                            Excel.Range Cell_nedv_in = excelworksheet.get_Range("L57");
                            string cell_nedv_in = Cell_nedv_in.Text.ToString();
                            decimal Nedv_in_yes = Convert.ToDecimal(cell_nedv_in);
                            excelcells = excelworksheet.get_Range("I57", Type.Missing);
                            if (vid_plat_pp == "Дебет")    //платят нам
                            {
                                excelcells.Value2 = Nedv_in_yes + pole_170;
                                pole_170 = Nedv_in_yes + pole_170;
                            }
                            if (vid_plat_pp == "Кредит")   //платим мы
                            {
                                excelcells.Value2 = Nedv_in_yes - pole_170;
                                pole_170 = Nedv_in_yes - pole_170;
                            }
                            #endregion
                            #region //Сумма активов
                            excelcells = excelworksheet.get_Range("I96", Type.Missing);
                            if (vid_plat_pp == "Дебет")
                            {
                                excelcells.Value2 = (pole_010) + (Nedv_rus_yes + pole_160) + (Nedv_in_yes + pole_170);
                                pole_270 = (pole_010) + (Nedv_rus_yes + pole_160) + (Nedv_in_yes + pole_170);
                            }
                            if (vid_plat_pp == "Кредит")
                            {
                                excelcells.Value2 = (pole_010) + (Nedv_rus_yes - pole_160) + (Nedv_in_yes - pole_170);
                                pole_270 = (pole_010) + (Nedv_rus_yes - pole_160) + (Nedv_in_yes - pole_170);
                            }
                            #endregion
                            #region //Кред. задолженность
                            excelcells = excelworksheet.get_Range("I100", Type.Missing);
                            excelcells.Value2 = pole_300;
                            #endregion
                            #region //Резерв на вознагр.
                            excelcells = excelworksheet.get_Range("I101", Type.Missing);
                            excelcells.Value2 = pole_310;
                            #endregion
                            #region //Сумма обязательств
                            excelcells = excelworksheet.get_Range("I105", Type.Missing);
                            excelcells.Value2 = pole_330;
                            #endregion
                            #region //СЧА
                            excelcells = excelworksheet.get_Range("I106", Type.Missing);
                            if (vid_plat_pp == "Дебет")
                            {
                                excelcells.Value2 = pole_270 - pole_330;
                                pole_400 = pole_270 - pole_330;
                            }
                            if (vid_plat_pp == "Кредит")
                            {
                                excelcells.Value2 = pole_270 - pole_330;
                                pole_400 = pole_270 - pole_330;
                            }
                            #endregion
                            #region //Кол. паев
                            excelcells = excelworksheet.get_Range("I107", Type.Missing);
                            excelcells.Value2 = pole_500;
                            #endregion
                            #region //Цена пая
                            excelcells = excelworksheet.get_Range("I109", Type.Missing);
                            if (vid_plat_pp == "Дебет")
                            {
                                excelcells.Value2 = pole_400 / Convert.ToDecimal(pole_500);
                            }
                            if (vid_plat_pp == "Кредит")
                            {
                                excelcells.Value2 = pole_400 / Convert.ToDecimal(pole_500);
                            }
                            #endregion

                            #endregion
                            //перезапись таблицы СЧА_день
                            //InsertData_Scha_day(data_v_boxe, (nomer_proxoda + 1));

                            #region //Сохранение файла
                            newfile_name = newfile_n + "\\" + nazv_otch + ".xls";
                            newfile_name_arh = path_to_arh_otch + "\\" + nazv_otch + ".xls";
                            //MessageBox.Show(newfile_name, "newfile_name");

                            objWorkExcel.DisplayAlerts = false;
                            objWorkBook.SaveAs(@newfile_name, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                            objWorkExcel.Quit();
                            File.Copy(newfile_name, newfile_name_arh, true);
                            #endregion

                            #region //постановка отметки об учете
                            string Update_PP = "UPDATE Данные_день_СЧА SET [Отметка_об_учете] = ? WHERE [Название_отчета] = ? AND [Номер_пп] = ?";
                            using (OleDbCommand CommandBDParams = new OleDbCommand(Update_PP, conn))
                            {
                                CommandBDParams.Parameters.Add("@Q1", OleDbType.Boolean).Value = true;
                                CommandBDParams.Parameters.Add("@Q2", OleDbType.Char).Value = nazv_otch;
                                CommandBDParams.Parameters.Add("@Q3", OleDbType.Integer).Value = nomer_PP;
                                CommandBDParams.ExecuteNonQuery();
                            }
                            #endregion
                        }
                        #endregion
                        kontrol_kol_strok++;
                    }
                    else
                    {
                        objWorkExcel.Quit();
                    }
                }
                #endregion
            }
        }

        protected internal void Zapis_v_excel_Scha_day_yesterday(string nazv_otcheta)
        {
            #region //получим директорию файла
            var obrazec_directory = new DirectoryInfo(@".\\Отчеты_на_отправку");
            string obrazec_name = "", newfile_n = "", newfile_name = "", newfile_name_arh = "";
            foreach (FileInfo file in obrazec_directory.GetFiles())
            {
                if (file.Name.Contains("Образец_сча_день"))
                { obrazec_name = file.FullName; newfile_n = file.DirectoryName; }
            }
            #endregion

            Excel.Sheets excelsheets;
            Excel.Range excelcells;
            Excel.Application objWorkExcel = new Excel.Application();

            string table_name = "SELECT * FROM СЧА_день";
            Connect_to_BD_for_otchet(table_name);
            var CommandBD = Connect_to_BD_for_otchet(table_name);
            var conn = CommandBD.Connection;
            OleDbDataReader dr1 = CommandBD.ExecuteReader();

            while (dr1.Read())
            {
                string name_otch_BD = dr1.GetString(0);
                DateTime date_otch = dr1.GetDateTime(2);

                #region //запись предыдущего дня
                string name_otch_yesterday = "СЧАд_" + UK + "_" + fond + "_за_" + Vybor_date_for_scha.vybran_date;
                if (date_otch.ToShortDateString() == Vybor_date_for_scha.vybran_date)
                {
                    #region //открыть образец
                    newfile_name2 = newfile_n + "\\" + nazv_otch + ".xls";
                    Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(obrazec_name, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    excelsheets = objWorkBook.Worksheets;                       //Получаем массив ссылок на листы выбранной книги
                    Excel.Worksheet excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);  //Получаем ссылку на лист 1
                    #endregion

                    #region //чтение предыдущих значений
                    string name_otch_y = dr1.GetString(0);
                    DateTime date_yesterday = dr1.GetDateTime(2);
                    decimal activ_all = dr1.GetDecimal(17);
                    decimal DC_all = dr1.GetDecimal(14);
                    decimal DC_rub = dr1.GetDecimal(15);
                    decimal DC_in = dr1.GetDecimal(16);
                    decimal obyz_all = dr1.GetDecimal(18);
                    decimal obyz_kred = dr1.GetDecimal(19);
                    decimal obyz_rez = dr1.GetDecimal(20);
                    decimal Scha = dr1.GetDecimal(7);
                    decimal nedv_rus = dr1.GetDecimal(21);
                    decimal nedv_in = dr1.GetDecimal(22);
                    decimal price_pay = dr1.GetDecimal(9);
                    double kol_p = dr1.GetDouble(8);
                    #endregion
                    #region //Дата предыдущая
                    excelcells = excelworksheet.get_Range("L14", Type.Missing);
                    excelcells.Value2 = Vybor_date_for_scha.vybran_date;
                    #endregion

                    #region //поля отчета

                    #region //ДС всего, рубли и валюта
                    excelcells = excelworksheet.get_Range("L16", Type.Missing);
                    excelcells.Value2 = DC_all;
                    excelcells = excelworksheet.get_Range("L18", Type.Missing);
                    excelcells.Value2 = DC_rub;
                    excelcells = excelworksheet.get_Range("L19", Type.Missing);
                    excelcells.Value2 = DC_in;
                    #endregion
                    #region //Недвиж. в РФ + вне РФ
                    excelcells = excelworksheet.get_Range("L53", Type.Missing);
                    excelcells.Value2 = nedv_rus;
                    excelcells = excelworksheet.get_Range("L57", Type.Missing);
                    excelcells.Value2 = nedv_in;
                    #endregion
                    #region //Сумма активов
                    excelcells = excelworksheet.get_Range("L96", Type.Missing);
                    excelcells.Value2 = activ_all;
                    #endregion
                    #region //Кред. задолженность
                    excelcells = excelworksheet.get_Range("L100", Type.Missing);
                    excelcells.Value2 = obyz_kred;
                    #endregion
                    #region //Резерв на вознагр.
                    excelcells = excelworksheet.get_Range("L101", Type.Missing);
                    excelcells.Value2 = obyz_rez;
                    #endregion
                    #region //Сумма обязательств
                    excelcells = excelworksheet.get_Range("L105", Type.Missing);
                    excelcells.Value2 = obyz_all;
                    #endregion
                    #region //СЧА
                    excelcells = excelworksheet.get_Range("L106", Type.Missing);
                    excelcells.Value2 = Scha;
                    #endregion
                    #region //Кол. паев
                    excelcells = excelworksheet.get_Range("L107", Type.Missing);
                    excelcells.Value2 = kol_p;
                    #endregion
                    #region //Цена пая
                    excelcells = excelworksheet.get_Range("L109", Type.Missing);
                    excelcells.Value2 = price_pay;
                    #endregion

                    #endregion

                    #region //Сохранение файла
                    newfile_name = newfile_n + "\\" + nazv_otcheta + ".xls";
                    newfile_name_arh = path_to_arh_otch + "\\" + nazv_otcheta + ".xls";
                    //MessageBox.Show(newfile_name, "newfile_name");

                    objWorkExcel.DisplayAlerts = false;
                    objWorkBook.SaveAs(@newfile_name, Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    objWorkExcel.Quit();
                    File.Copy(newfile_name, newfile_name_arh, true);
                    #endregion
                }
                else
                {
                    objWorkExcel.Quit();
                }
                #endregion
            }
        }

        protected internal void Zapis_v_excel_Scha_day_last(string nazv_otcheta)
        {
            #region //получим директорию файла
            var obrazec_directory = new DirectoryInfo(@".\\Отчеты_на_отправку");
            string obrazec_name = "", newfile_n = "", newfile_name = "", newfile_name_arh = "";
            foreach (FileInfo file in obrazec_directory.GetFiles())
            {
                if (file.Name.Contains("Образец_сча_день"))
                { obrazec_name = file.FullName; newfile_n = file.DirectoryName; }
            }
            #endregion

            Excel.Sheets excelsheets;
            Excel.Range excelcells;
            Excel.Application objWorkExcel = new Excel.Application();

            string text_connecta = "SELECT * FROM Платежные_поручения";
            Connect_to_BD_for_otchet(text_connecta);
            var CommandBD = Connect_to_BD_for_otchet(text_connecta);
            var conn = CommandBD.Connection;

            newfile_name2 = newfile_n + "\\" + nazv_otch + ".xls";
            Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(newfile_name2, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            excelsheets = objWorkBook.Worksheets;                       //Получаем массив ссылок на листы выбранной книги
            Excel.Worksheet excelworksheet = (Excel.Worksheet)excelsheets.get_Item(1);  //Получаем ссылку на лист 1

            #region //ДС всего, руб + вал
            Excel.Range Cell_DC_all = excelworksheet.get_Range("I16");
            string cell_dc_all = Cell_DC_all.Text.ToString();
            decimal DC_all = Convert.ToDecimal(cell_dc_all);
            Excel.Range Cell_DC_rub = excelworksheet.get_Range("I18");
            string cell_dc_rub = Cell_DC_rub.Text.ToString();
            decimal DC_rub = Convert.ToDecimal(cell_dc_rub);
            Excel.Range Cell_DC_val = excelworksheet.get_Range("I19");
            string cell_dc_val = Cell_DC_val.Text.ToString();
            decimal DC_val = Convert.ToDecimal(cell_dc_val);
            #endregion
            #region //Недвиж. РФ и не РФ
            Excel.Range Cell_ned_rus = excelworksheet.get_Range("I53");
            string cell_ned_rus = Cell_ned_rus.Text.ToString();
            decimal Ned_rus = Convert.ToDecimal(cell_ned_rus);
            Excel.Range Cell_ned_ne_rus = excelworksheet.get_Range("I57");
            string cell_ned_ne_rus = Cell_ned_ne_rus.Text.ToString();
            decimal Ned_ne_rus = Convert.ToDecimal(cell_ned_ne_rus);
            #endregion
            #region //Сумма активов
            Excel.Range Cell_sum_act = excelworksheet.get_Range("I96");
            string cell_sum_act = Cell_sum_act.Text.ToString();
            decimal Sum_act = Convert.ToDecimal(cell_sum_act);
            #endregion
            #region //Кред задолж
            Excel.Range Cell_kred_z = excelworksheet.get_Range("I100");
            string cell_kred_z = Cell_kred_z.Text.ToString();
            decimal kred_z = Convert.ToDecimal(cell_kred_z);
            #endregion
            #region //Резерв
            Excel.Range Cell_rez = excelworksheet.get_Range("I101");
            string cell_rez = Cell_rez.Text.ToString();
            decimal rezerv = Convert.ToDecimal(cell_rez);
            #endregion
            #region //Сумма обяз
            Excel.Range Cell_sum_obyz = excelworksheet.get_Range("I105");
            string cell_sum_obyz = Cell_sum_obyz.Text.ToString();
            decimal Sum_obyz = Convert.ToDecimal(cell_sum_obyz);
            #endregion
            #region //СЧА
            Excel.Range Cell_scha = excelworksheet.get_Range("I106");
            string cell_scha = Cell_scha.Text.ToString();
            decimal Scha = Convert.ToDecimal(cell_scha);
            #endregion
            #region //Кол паев
            Excel.Range Cell_kol_p = excelworksheet.get_Range("I107");
            string cell_kol_p = Cell_kol_p.Text.ToString();
            double Kol_P = Convert.ToDouble(cell_kol_p);
            #endregion
            #region //Цена пая
            Excel.Range Cell_price_pay = excelworksheet.get_Range("I109");
            string cell_price_pay = Cell_price_pay.Text.ToString();
            decimal Price_pay = Convert.ToDecimal(cell_price_pay);
            #endregion

            #region //запись в БД
            string Update_vipiska_day_last = "UPDATE СЧА_день SET [СЧА] = ?, [Кол_паев] = ?, [Цена_пая] = ?, [ДС_всего] = ?, [ДС_руб] = ?, [ДС_ин] = ?, [Актив_всего] = ?, [Обяз_всего] = ?, [Обяз_кред_задолжн] = ?, [Обяз_резерв] = ?, [Недвиж_РФ] = ?, [Недвиж_не_РФ] = ? WHERE [Название_отчета] = ?";
            using (OleDbCommand CommandBDParams = new OleDbCommand(Update_vipiska_day_last, conn))
            {
                CommandBDParams.Parameters.Add("@Q1", OleDbType.Decimal).Value = Scha;
                CommandBDParams.Parameters.Add("@Q2", OleDbType.Double).Value = Kol_P;
                CommandBDParams.Parameters.Add("@Q3", OleDbType.Double).Value = Price_pay;
                CommandBDParams.Parameters.Add("@Q4", OleDbType.Decimal).Value = DC_all;
                CommandBDParams.Parameters.Add("@Q5", OleDbType.Decimal).Value = DC_rub;
                CommandBDParams.Parameters.Add("@Q6", OleDbType.Decimal).Value = DC_val;
                CommandBDParams.Parameters.Add("@Q7", OleDbType.Decimal).Value = Sum_act;
                CommandBDParams.Parameters.Add("@Q8", OleDbType.Decimal).Value = Sum_obyz;
                CommandBDParams.Parameters.Add("@Q9", OleDbType.Decimal).Value = kred_z;
                CommandBDParams.Parameters.Add("@Q10", OleDbType.Decimal).Value = rezerv;
                CommandBDParams.Parameters.Add("@Q11", OleDbType.Decimal).Value = Ned_rus;
                CommandBDParams.Parameters.Add("@Q12", OleDbType.Decimal).Value = Ned_ne_rus;
                CommandBDParams.Parameters.Add("@Q3", OleDbType.Char).Value = nazv_otcheta;
                CommandBDParams.ExecuteNonQuery();
            }
            #endregion

            objWorkBook.Close();
            objWorkExcel.Quit();
        }
        protected internal void Make_otchet()       //метод формирования отчетов
        {
            string vvod_data = work_form.calendar1.Selected_date;
            string vvod_fonda = work_form.comboBox_Fond_otch.Text;
            string vvod_UK = work_form.comboBox_UK_otch.Text;
            string vvod_type_otchet = work_form.comboBox_type_otch.Text;
            int kontrol_vvoda = 2;

            #region Контроль ввода всех параметров
            if (string.IsNullOrWhiteSpace(vvod_data))
            {
                MessageBox.Show("Ошибка! Выберите дату!", "Уведомление", MessageBoxButtons.OK);
                kontrol_vvoda = 1;
            }
            else
            {
                if (string.IsNullOrWhiteSpace(vvod_fonda))
                {
                    MessageBox.Show("Ошибка! Выберите фонд!", "Уведомление", MessageBoxButtons.OK);
                    kontrol_vvoda = 1;
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(vvod_UK))
                    {
                        MessageBox.Show("Ошибка! Выберите Управляющую компанию!", "Уведомление", MessageBoxButtons.OK);
                        kontrol_vvoda = 1;
                    }
                    else
                    {
                        if (string.IsNullOrWhiteSpace(vvod_type_otchet))
                        {
                            MessageBox.Show("Ошибка! Выберите тип формируемого отчета!", "Уведомление", MessageBoxButtons.OK);
                            kontrol_vvoda = 1;
                        }
                        else
                        {
                            kontrol_vvoda = 0;
                        }
                    }
                }
            }
            #endregion

            if (kontrol_vvoda == 0)
            {
                #region Выписка за день
                if (vvod_type_otchet == "Выписка за день")
                {
                    //показ формы для ввода валюты
                    {
                        DialogResult vvod_val = MessageBox.Show("Следует ввести валюту и курс?" + Environment.NewLine + "Без выбора - расчет в рублях!", "Уведомление", MessageBoxButtons.YesNo);
                        if (vvod_val == DialogResult.Yes) { Form Vvod_kursa = new Vvod_kursa(); Vvod_kursa.ShowDialog(); }
                        if (vvod_val == DialogResult.No) { Vvod_kursa.vybran_val = "RUB"; Vvod_kursa.vvod_kurs = 1; }
                    }
                    if (Vvod_kursa.cancel_kod == true) { return; }

                    //подсчет строк в таблице
                    string text_podkl_chet = "SELECT count(*) FROM Платежные_поручения";
                    Connect_to_BD_for_otchet(text_podkl_chet);
                    var CommandBD_prob = Connect_to_BD_for_otchet(text_podkl_chet);
                    var conn_prob = CommandBD_prob.Connection;
                    int kol_strok = Convert.ToInt32(CommandBD_prob.ExecuteScalar());

                    int n = 0, n2 = 1, kontrol_data = 0;
                    DateTime data_iz_BD = new DateTime();

                    //проверка
                    //MessageBox.Show(Convert.ToString(kol_strok), "кол строк в табл п/п");
                    //MessageBox.Show(Convert.ToString(vvod_data), "дата выбранная в календаре");

                    //CommandBD.CommandText = "SELECT * FROM Платежные_поручения";                     //выбор всего из таблицы 
                    string text_podkl = "SELECT * FROM Платежные_поручения";
                    Connect_to_BD_for_otchet(text_podkl);
                    var CommandBD = Connect_to_BD_for_otchet(text_podkl);
                    var conn = CommandBD.Connection;
                    OleDbDataReader dr1 = CommandBD.ExecuteReader();

                    while (dr1.Read())
                    {
                        data_iz_BD = dr1.GetDateTime(4);
                        data_plat_BD = Convert.ToDateTime(data_iz_BD).ToShortDateString();
                        vid_plat = dr1.GetString(9);
                        otmetka = dr1.GetBoolean(12);
                        if (otmetka == false)
                        {
                            if (data_plat_BD == vvod_data)                                    //сравнение в БД дат
                            {
                                nomer_PP = dr1.GetInt32(0);     //номер п/п
                                UK = dr1.GetString(1);          //название УК
                                fond = dr1.GetString(2);        //название фонда
                                agent = dr1.GetString(3);       //название контрагента
                                valuta = dr1.GetString(5);      //валюта
                                summa = dr1.GetDecimal(6);      //сумма 
                                naznach = dr1.GetString(15);    //назначение

                                if (valuta == Vvod_kursa.vybran_val)
                                {
                                    if (valuta == "RUB")
                                    {
                                        #region
                                        if (vid_plat == "Дебет")
                                        {
                                            summa_deb = summa;          //сумма ушла в дебет
                                            oborot_deb = oborot_deb + summa;
                                        }
                                        else if (vid_plat == "Кредит")
                                        {
                                            summa_kred = summa;         //сумма ушла в кредит
                                            oborot_kred = oborot_kred + summa;
                                        }
                                        else
                                        {
                                            MessageBox.Show("Ошибка в БД! Исправьте вид платежа!", "Уведомление", MessageBoxButtons.OK);
                                        }

                                        #region //вытаскиваем СЧЕТ УК
                                        string text_podkl_chetUK = "SELECT * FROM Список_Управляющих_компаний";
                                        Connect_to_BD_for_otchet(text_podkl_chetUK);
                                        var CommandBD2 = Connect_to_BD_for_otchet(text_podkl_chetUK);
                                        var conn2 = CommandBD2.Connection;
                                        OleDbDataReader dr2 = CommandBD2.ExecuteReader();
                                        while (n2 < 200)
                                        {
                                            dr2.Read();                                                  //чтение пошло
                                            UK_BD = dr2.GetString(3);
                                            if (UK == UK_BD)                                    //сравнение названия УК в БД
                                            {
                                                //счет УК
                                                chet_UK = dr2.GetString(10);
                                                //входящее сальдо кредит
                                                vxod_saldo_kred = dr2.GetDecimal(11);
                                                n2 = n2 + 500;
                                            }
                                            n2++;
                                        }
                                        dr2.Close();
                                        n2 = 1;
                                        #endregion
                                        string data_bez_time = data_plat_BD.Substring(0, 10);
                                        //название выписки
                                        nazv_vipiski = "Выписка_"+ valuta + "_" + UK + "_" + fond + "_за_" + data_bez_time;

                                        //исходящие сальдо кредит и дебет
                                        isxod_saldo_kred = vxod_saldo_kred + summa_kred - summa_deb;              //кредит
                                        if (isxod_saldo_kred < 0)
                                        {
                                            isxod_saldo_deb = isxod_saldo_kred * (-1);                             //дебет
                                        }

                                        #region
                                        /*MessageBox.Show(vvod_data, "дата календаря");
                                        MessageBox.Show(Convert.ToString(data_iz_BD), "дата из БД");
                                        MessageBox.Show(data_plat_BD, "дата п/п из БД");
                                        MessageBox.Show(vid_plat);
                                        MessageBox.Show(Convert.ToString(nomer_PP));
                                        MessageBox.Show(UK);
                                        MessageBox.Show(chet_UK);
                                        MessageBox.Show(fond);
                                        MessageBox.Show(valuta);
                                        MessageBox.Show(agent);
                                        MessageBox.Show(naznach);
                                        MessageBox.Show(nazv_vipiski);
                                        */
                                        #endregion
                                        nomer_iteracii++;
                                        InsertData_vyp(vvod_data, nomer_iteracii);                            //вызов метода с верху (запись/обновление БД)

                                        Zapis_v_excel_vipiska_day(nazv_vipiski);                        //Начало создания файла excel

                                        kontrol_data = kontrol_data + 1000;                //таким образом не будет сообщения об отсутствии операций
                                        #endregion
                                    }
                                    else
                                    {
                                        #region
                                        if (vid_plat == "Дебет")
                                        {
                                            summa_deb = summa * Vvod_kursa.vvod_kurs;          //сумма ушла в дебет
                                            oborot_deb = oborot_deb + summa * Vvod_kursa.vvod_kurs;
                                        }
                                        else if (vid_plat == "Кредит")
                                        {
                                            summa_kred = summa * Vvod_kursa.vvod_kurs;         //сумма ушла в кредит
                                            oborot_kred = oborot_kred + summa * Vvod_kursa.vvod_kurs;
                                        }
                                        else
                                        {
                                            MessageBox.Show("Ошибка в БД! Исправьте вид платежа!", "Уведомление", MessageBoxButtons.OK);
                                        }

                                        #region //вытаскиваем СЧЕТ УК
                                        string text_podkl_chetUK = "SELECT * FROM Список_Управляющих_компаний";
                                        Connect_to_BD_for_otchet(text_podkl_chetUK);
                                        var CommandBD2 = Connect_to_BD_for_otchet(text_podkl_chetUK);
                                        var conn2 = CommandBD2.Connection;
                                        OleDbDataReader dr2 = CommandBD2.ExecuteReader();
                                        while (n2 < 200)
                                        {
                                            dr2.Read();                                                  //чтение пошло
                                            UK_BD = dr2.GetString(3);
                                            if (UK == UK_BD)                                    //сравнение названия УК в БД
                                            {
                                                //счет УК
                                                chet_UK = dr2.GetString(10);
                                                //входящее сальдо кредит
                                                vxod_saldo_kred = dr2.GetDecimal(11);
                                                n2 = n2 + 500;
                                            }
                                            n2++;
                                        }
                                        dr2.Close();
                                        n2 = 1;
                                        #endregion
                                        string data_bez_time = data_plat_BD.Substring(0, 10);
                                        //название выписки
                                        nazv_vipiski = "Выписка_" + valuta + "_" + UK + "_" + fond + "_за_" + data_bez_time;

                                        //исходящие сальдо кредит и дебет
                                        isxod_saldo_kred = vxod_saldo_kred + summa_kred - summa_deb;              //кредит
                                        if (isxod_saldo_kred < 0)
                                        {
                                            isxod_saldo_deb = isxod_saldo_kred * (-1);                             //дебет
                                        }
                                        
                                        nomer_iteracii++;
                                        InsertData_vyp(vvod_data, nomer_iteracii);                            //вызов метода с верху (запись/обновление БД)

                                        Zapis_v_excel_vipiska_day(nazv_vipiski);                        //Начало создания файла excel

                                        kontrol_data = kontrol_data + 1000;                //таким образом не будет сообщения об отсутствии операций
                                        #endregion
                                    }
                                }
                                else
                                {
                                    kontrol_data = kontrol_data - 1;                    //таким образом Будут сообщения об отсутствии операций
                                    n++;
                                }
                            }
                            else
                            {
                                kontrol_data = kontrol_data - 1;                    //таким образом Будут сообщения об отсутствии операций
                                n++;
                            }
                            
                        }
                    }
                    dr1.Close();

                    if (kontrol_data > 0)
                    {
                        MessageBox.Show("Отчет сформирован!", "Уведомление", MessageBoxButtons.OK);
                    }
                    if (kontrol_data < 0)
                    {
                        MessageBox.Show("За данную дату операций не происходило!" + Environment.NewLine + "Следует создать пустой отчет вручную!", "Уведомление", MessageBoxButtons.OK);
                    }
                }
                #endregion

                #region СЧА день
                if (vvod_type_otchet == "Стоимость чистых активов за день")
                {
                    //показ формы для ввода валюты
                    {
                        DialogResult vvod_val = MessageBox.Show("Следует ввести валюту и курс?", "Уведомление", MessageBoxButtons.YesNo);
                        if (vvod_val == DialogResult.Yes) { Form Vvod_kursa = new Vvod_kursa(); Vvod_kursa.ShowDialog(); }
                        if (vvod_val == DialogResult.No) { Vvod_kursa.vvod_kurs = 1; }
                    }
                    if (Vvod_kursa.cancel_kod == true) { return; }

                    //показ формы для ввода предыдущей даты
                    { Form Vybor_date = new Vybor_date_for_scha(); Vybor_date.ShowDialog(); }
                    if (Vybor_date_for_scha.cancel_kod_date == true) { return; }

                    #region //зануление
                    pole_010 = 0; pole_011 = 0; pole_012 = 0; pole_160 = 0; pole_300 = 0; pole_310 = 0; pole_270 = 0; pole_330 = 0; pole_400 = 0; pole_600 = 0; pole_170 = 0;
                    #endregion

                    #region //подсчет строк в таблице
                    string text_podkl_chet = "SELECT count(*) FROM Платежные_поручения";
                    Connect_to_BD_for_otchet(text_podkl_chet);
                    var CommandBD_prob = Connect_to_BD_for_otchet(text_podkl_chet);
                    var conn_prob = CommandBD_prob.Connection;
                    int kol_strok = Convert.ToInt32(CommandBD_prob.ExecuteScalar());
                    #endregion
                    DateTime data_iz_BD = new DateTime();
                    int kontrol_data = 0;

                    string text_podkl = "SELECT * FROM Платежные_поручения";
                    Connect_to_BD_for_otchet(text_podkl);
                    var CommandBD = Connect_to_BD_for_otchet(text_podkl);
                    var conn = CommandBD.Connection;
                    OleDbDataReader dr1 = CommandBD.ExecuteReader();
                    while (dr1.Read())
                    {
                        data_iz_BD = dr1.GetDateTime(4);
                        data_plat_BD = Convert.ToDateTime(data_iz_BD).ToShortDateString();
                        vid_plat = dr1.GetString(9);    //вид платежа
                        otmetka = dr1.GetBoolean(16);
                        if (otmetka == false)
                        {
                            if (data_plat_BD == vvod_data)                                    //сравнение в БД дат
                            {
                                #region //данные п/п
                                nomer_PP = dr1.GetInt32(0);     //номер п/п
                                UK = dr1.GetString(1);          //название УК
                                fond = dr1.GetString(2);        //название фонда
                                agent = dr1.GetString(3);       //название контрагента
                                valuta = dr1.GetString(5);      //валюта
                                summa = dr1.GetDecimal(6);      //сумма 
                                naznach = dr1.GetString(15);    //назначение
                                type_raspor = dr1.GetString(14); //тип распоряжения
                                kol_pay = dr1.GetDouble(10);    //кол паев
                                #endregion

                                #region //кол паев
                                string text_podkl_pay_UK = "SELECT * FROM Список_Управляющих_компаний";
                                Connect_to_BD_for_otchet(text_podkl_pay_UK);
                                var CommandBD2 = Connect_to_BD_for_otchet(text_podkl_pay_UK);
                                var conn2 = CommandBD2.Connection;
                                OleDbDataReader dr2 = CommandBD2.ExecuteReader();
                                int n2 = 0;
                                while (n2 < 200)
                                {
                                    dr2.Read();                                                  //чтение пошло
                                    UK_BD = dr2.GetString(3);
                                    if (UK == UK_BD)                                    //сравнение названия УК в БД
                                    {
                                        //текущие паи УК
                                        pay_UK_now = dr2.GetDouble(12);
                                        //лицензия УК
                                        nomer_lic_UK = dr2.GetString(5);
                                        n2 = n2 + 500;
                                    }
                                    n2++;
                                }
                                dr2.Close();
                                n2 = 1;
                                #endregion
                                #region //лицензия фонда
                                string text_podkl_to_fond = "SELECT * FROM Список_фондов";
                                Connect_to_BD_for_otchet(text_podkl_to_fond);
                                var CommandBD3 = Connect_to_BD_for_otchet(text_podkl_to_fond);
                                var conn3 = CommandBD3.Connection;
                                OleDbDataReader dr3 = CommandBD3.ExecuteReader();
                                int n3 = 0;
                                while (n3 < 200)
                                {
                                    dr3.Read();                                                  //чтение пошло
                                    Fond_BD = dr3.GetString(2);
                                    if (fond == Fond_BD)                                    //сравнение названия УК в БД
                                    {
                                        //лицензия
                                        nomer_lic_fond = dr3.GetString(3);
                                        n3 = n3 + 500;
                                    }
                                    n3++;
                                }
                                dr3.Close();
                                n3 = 1;
                                #endregion

                                #region //переводы
                                if (type_raspor == "Перевод")
                                {
                                    if (vid_plat == "Дебет")    //платят нам
                                    {
                                        if (valuta == "RUB")
                                        {
                                            pole_011 = pole_011 + summa;
                                        }
                                        else
                                        {
                                            pole_012 = pole_012 + summa;
                                        }
                                    }
                                    if (vid_plat == "Кредит")   //платим мы
                                    {
                                        if (valuta == "RUB")
                                        {
                                            pole_011 = pole_011 - summa;
                                        }
                                        else
                                        {
                                            pole_012 = pole_012 - summa;
                                        }
                                    }
                                }
                                #endregion
                                #region //оплата услуг
                                if (type_raspor == "Оплата услуг")
                                {
                                    if (naznach.Contains("депозитар")) { kontrol_oplat_spec_dep = 1; }
                                    if (naznach.Contains("регистр")) { kontrol_oplat_spec_reg = 1; }
                                    if (kontrol_oplat_spec_dep == 1)
                                    {
                                        if (vid_plat == "Дебет")    //платят нам
                                        {
                                            if (valuta == "RUB")
                                            {
                                                pole_011 = pole_011 + summa;
                                            }
                                            else
                                            {
                                                pole_012 = pole_012 + summa;
                                            }
                                        }
                                        if (vid_plat == "Кредит")   //платим мы
                                        {
                                            if (valuta == "RUB")
                                            {
                                                pole_011 = pole_011 - summa;
                                            }
                                            else
                                            {
                                                pole_012 = pole_012 - summa;
                                            }
                                        }
                                    }
                                    else { pole_310 = pole_310 + summa; }
                                    if (kontrol_oplat_spec_reg == 1)
                                    {
                                        if (vid_plat == "Дебет")    //платят нам
                                        {
                                            if (valuta == "RUB")
                                            {
                                                pole_011 = pole_011 + summa;
                                            }
                                            else
                                            {
                                                pole_012 = pole_012 + summa;
                                            }
                                        }
                                        if (vid_plat == "Кредит")   //платим мы
                                        {
                                            if (valuta == "RUB")
                                            {
                                                pole_011 = pole_011 - summa;
                                            }
                                            else
                                            {
                                                pole_012 = pole_012 - summa;
                                            }

                                        }
                                    }
                                    else { pole_310 = pole_310 + summa; }
                                }
                                #endregion
                                #region //вознаграждения
                                if (type_raspor == "Вознаграждение")
                                {
                                    if (vid_plat == "Дебет")    //платят нам
                                    {
                                        if (valuta == "RUB")
                                        {
                                            pole_011 = pole_011 + summa;
                                        }
                                        else
                                        {
                                            pole_012 = pole_012 + summa;
                                        }
                                    }
                                    if (vid_plat == "Кредит")   //платим мы
                                    {
                                        if (valuta == "RUB")
                                        {
                                            pole_011 = pole_011 - summa;
                                        }
                                        else
                                        {
                                            pole_012 = pole_012 - summa;
                                        }
                                    }
                                }
                                #endregion
                                #region //покупка
                                if (type_raspor == "Покупка")
                                {
                                    if (vid_plat == "Кредит")   //платим мы
                                    {
                                        if (valuta == "RUB")
                                        {
                                            pole_011 = pole_011 - summa;
                                        }
                                        else
                                        {
                                            pole_012 = pole_012 - summa;
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Проверьте вид платежа у плпт поручения №" + nomer_PP + " !" + Environment.NewLine + "Возможно потребуется исправление отчета!", "Уведомление");
                                    }
                                }
                                #endregion
                                #region //продажа
                                if (type_raspor == "Продажа")
                                {
                                    if (vid_plat == "Дебет")   //платят нам
                                    {
                                        if (valuta == "RUB")
                                        {
                                            pole_011 = pole_011 + summa;
                                        }
                                        else
                                        {
                                            pole_012 = pole_012 + summa;
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Проверьте вид платежа у плпт поручения №" + nomer_PP + " !" + Environment.NewLine + "Возможно потребуется исправление отчета!", "Уведомление");
                                    }
                                }
                                #endregion
                                #region //покупка паев нам (погашение паев)
                                if (type_raspor == "Продажа паев")
                                {
                                    if (valuta == "RUB")
                                    {
                                        pole_011 = pole_011 - summa;
                                    }
                                    else
                                    {
                                        pole_012 = pole_012 - summa;
                                    }
                                    pole_500 = pay_UK_now - kol_pay;
                                }
                                #endregion
                                #region //продажа паев нами
                                if (type_raspor == "Покупка паев")
                                {
                                    if (valuta == "RUB")
                                    {
                                        pole_011 = pole_011 + summa;
                                    }
                                    else
                                    {
                                        pole_012 = pole_012 + summa;
                                    }
                                    pole_500 = pay_UK_now + kol_pay;
                                }
                                #endregion
                                #region //паи без изменения
                                if ((type_raspor != "Продажа паев") && (type_raspor != "Покупка паев")) { pole_500 = pay_UK_now; }
                                #endregion
                                #region //покупка недвижимости
                                if (type_raspor == "Покупка недвижимости")
                                {
                                    if (vid_plat == "Кредит")   //платим мы
                                    {
                                        if (valuta == "RUB")
                                        {
                                            pole_160 = pole_160 - summa;
                                        }
                                        else
                                        {
                                            pole_170 = pole_170 - summa;
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Проверьте вид платежа у плпт поручения №" + nomer_PP + " !" + Environment.NewLine + "Возможно потребуется исправление отчета!", "Уведомление");
                                    }
                                }
                                #endregion
                                #region //продажа недвижимости
                                if (type_raspor == "Продажа недвижимости")
                                {
                                    if (vid_plat == "Дебет")   //платят нам
                                    {
                                        if (valuta == "RUB")
                                        {
                                            pole_160 = pole_160 + summa;
                                        }
                                        else
                                        {
                                            pole_170 = pole_170 + summa;
                                        }
                                    }
                                    else
                                    {
                                        MessageBox.Show("Проверьте вид платежа у плпт поручения №" + nomer_PP + " !" + Environment.NewLine + "Возможно потребуется исправление отчета!", "Уведомление");
                                    }
                                }
                                #endregion
                                #region //ДС на счете
                                pole_010 = pole_011 + (pole_012 * Vvod_kursa.vvod_kurs);
                                if (pole_010 < 0)
                                {
                                    if (pole_011 < 0)
                                    {
                                        if (pole_012 < 0)
                                        {
                                            pole_300 = pole_300 + (pole_011 * (-1)) + (pole_012 * (-1) * Vvod_kursa.vvod_kurs);
                                            pole_011 = 0;
                                            pole_012 = 0;
                                        }
                                        else
                                        {
                                            pole_300 = pole_300 + (pole_011 * (-1)) + (pole_012 * Vvod_kursa.vvod_kurs);
                                            pole_011 = 0;
                                            pole_012 = 0;
                                        }
                                    }
                                    else if (pole_012 < 0)
                                    {
                                        if (pole_011 < 0)
                                        {
                                            pole_300 = pole_300 + (pole_011 * (-1)) + (pole_012 * (-1) * Vvod_kursa.vvod_kurs);
                                            pole_011 = 0;
                                            pole_012 = 0;
                                        }
                                        else
                                        {
                                            pole_300 = pole_300 + pole_011 + (pole_012 * (-1) * Vvod_kursa.vvod_kurs);
                                            pole_011 = 0;
                                            pole_012 = 0;
                                        }
                                    }
                                    pole_010 = 0;
                                }
                                #endregion

                                if (pole_160 < 0) { pole_300 = pole_300 + (pole_160 * (-1)); }
                                if (pole_170 < 0) { pole_300 = pole_300 + (pole_170 * (-1)); }

                                //сумма активов
                                pole_270 = pole_010 + pole_160 + pole_170;
                                //сумма обязательств
                                pole_330 = pole_300 + pole_310;
                                //сча
                                pole_400 = pole_270 - pole_330;
                                //кол паев
                                pole_500 = pole_500;
                                //расчетная стоимость пая
                                decimal kontrol = Convert.ToDecimal(pole_500);
                                if (kontrol != 0) { pole_600 = Math.Round(pole_400 / kontrol, 2, MidpointRounding.AwayFromZero); }
                                else { pole_600 = Math.Round(pole_400, 2, MidpointRounding.AwayFromZero); }

                                #region //запись в БД + создание excel
                                //название отчета
                                string data_bez_time = vvod_data.Substring(0, 10);
                                nazv_otch = "СЧАд_" + UK + "_" + fond + "_за_" + data_bez_time;
                                //запись в БД
                                InsertData_Scha_day(vvod_data, nomer_iteracii);
                                //запись данных предыдущего дня
                                Zapis_v_excel_Scha_day_yesterday(nazv_otch);
                                //создание excel текущего дня
                                Zapis_v_excel_Scha_day_today(nazv_otch, vid_plat, pole_010, pole_011,pole_012,pole_160,pole_300,pole_310,pole_270,pole_330,pole_400,pole_600,pole_170);
                                //последнее обновление таблиц БД
                                Zapis_v_excel_Scha_day_last(nazv_otch);
                                #endregion

                                nomer_iteracii++;
                                kontrol_data = kontrol_data + 1000;
                            }
                            else
                            {
                                kontrol_data = kontrol_data - 1;                    //таким образом Будут сообщения об отсутствии операций
                            }
                        }
                    }
                    dr1.Close();

                    if (kontrol_data > 0)
                    {
                        MessageBox.Show("Отчет сформирован!", "Уведомление", MessageBoxButtons.OK);
                    }
                    if (kontrol_data < 0)
                    {
                        MessageBox.Show("За данную дату операций не происходило!" + Environment.NewLine + "Следует создать пустой отчет вручную!", "Уведомление", MessageBoxButtons.OK);
                    }
                }
                #endregion

                #region СЧА месяц
                if (vvod_type_otchet == "Стоимость чистых активов за месяц")
                {

                }
                #endregion

                #region Стоимость имущества
                if (vvod_type_otchet == "Стоимость имущества")
                {

                }
                #endregion

                #region Владельцы паев
                if (vvod_type_otchet == "Владельцы акций/паев")
                {

                }
                #endregion

            }
        }
        #endregion

        #region Просмотр отчетов
        protected internal void Watch_otchet()
        {
            try
            {
                string vvod_data = work_form.calendar1.Selected_date;
                string vvod_fonda = work_form.comboBox_Fond_otch.Text;
                string vvod_UK = work_form.comboBox_UK_otch.Text;
                string vvod_type_otchet = work_form.comboBox_type_otch.Text;
                int kontrol_vvoda = 2;

                #region Контроль ввода всех параметров
                if (string.IsNullOrWhiteSpace(vvod_data))
                {
                    MessageBox.Show("Ошибка! Выберите дату!", "Уведомление", MessageBoxButtons.OK);
                    kontrol_vvoda = 1;
                }
                else
                {
                    if (string.IsNullOrWhiteSpace(vvod_fonda))
                    {
                        MessageBox.Show("Ошибка! Выберите фонд!", "Уведомление", MessageBoxButtons.OK);
                        kontrol_vvoda = 1;
                    }
                    else
                    {
                        if (string.IsNullOrWhiteSpace(vvod_UK))
                        {
                            MessageBox.Show("Ошибка! Выберите Управляющую компанию!", "Уведомление", MessageBoxButtons.OK);
                            kontrol_vvoda = 1;
                        }
                        else
                        {
                            if (string.IsNullOrWhiteSpace(vvod_type_otchet))
                            {
                                MessageBox.Show("Ошибка! Выберите тип формируемого отчета!", "Уведомление", MessageBoxButtons.OK);
                                kontrol_vvoda = 1;
                            }
                            else
                            {
                                kontrol_vvoda = 0;
                            }
                        }
                    }
                }
                #endregion

                if (kontrol_vvoda == 0)
                {
                    string type_otchet = "";
                    #region Типы отчетов
                    if (vvod_type_otchet == "Выписка за день")
                    {
                        type_otchet = "Выписка";
                    }
                    if (vvod_type_otchet == "Стоимость чистых активов за день")
                    {
                        type_otchet = "СЧАд";
                    }
                    if (vvod_type_otchet == "Стоимость чистых активов за месяц")
                    {
                        type_otchet = "СЧАм";
                    }
                    if (vvod_type_otchet == "Стоимость имущества")
                    {
                        type_otchet = "Имущество";
                    }
                    if (vvod_type_otchet == "Владельцы акций/паев")
                    {
                        type_otchet = "Владельцы";
                    }
                    #endregion

                    string file_name_not_full = type_otchet + "_" + vvod_UK + "_" + vvod_fonda + "_за_" + vvod_data + ".xls";
                    string file_name_full = path_to_arh_otch + "\\" + file_name_not_full;

                    Process.Start(file_name_full);
                }
            }
            catch
            {
                MessageBox.Show("Ошибка");
            }
        }
        #endregion

        #region загрузка документов в бд

        /// <summary>
        /// Метод создания и записи в файл txt
        /// </summary>
        protected internal void Zapis_text(string put, string text_for_zap)
        {
            StreamWriter new_txt = File.AppendText(put);
            new_txt.WriteLine(text_for_zap);
            new_txt.Close();
        }
        
        /// <summary>
        /// Метод загрузки плат поручения с паями
        /// </summary>
        protected internal void Zagruzka_paychika(string Put_file, int max_strok_prov, int ID_fond, string Fond, string UK,
            string Shet, DateTime date, decimal kol_p, string type_oper)
        {
            string type_der;
            double kol_pay_all = new double();
            double kol_pay = Convert.ToDouble(kol_p);

            #region            //Подключение к БД
            OleDbCommand CommandBD_zap = new OleDbCommand();                                      //команда, через которую все делается
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DataBaseMy.accdb";
            OleDbConnection conn_zap = new OleDbConnection(connectionString);                       //новое подключение к БД
            CommandBD_zap.Connection = conn_zap;                                                      //соединение с бд
            conn_zap.Open();
            #endregion
            #region //Подключение Excel
            Excel.Application objWorkExcel = new Excel.Application();                   //подключим excel
            Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(Put_file, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
            Excel.Worksheet objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1];
            #endregion
            #region //проверка наличия пайщика
            string nameKontrBD;
            bool nalich_famil = false;
            int kontroz_nalichiy = 10, str_kontr = 0;

            Excel.Range CellKontr = objWorkSheet.get_Range("A25");
            string nameKontr = CellKontr.Text.ToString();
            CommandBD_zap.CommandText = "SELECT * FROM Пайщики_фонда";
            OleDbDataReader dr5 = CommandBD_zap.ExecuteReader();
            while (dr5.Read())
            {
                nameKontrBD = dr5.GetString(3);
                if (nameKontr != nameKontrBD)
                {
                    kontroz_nalichiy = 1; //нету совпадения
                    str_kontr++;
                }
                else if (nameKontr == nameKontrBD)
                {
                    kontroz_nalichiy = 0; //есть совпадение
                    kol_pay_all = dr5.GetDouble(6);
                    str_kontr = 0;
                    break;
                }
                if (str_kontr > max_strok_prov) { break; }
            }
            dr5.Close();
            #endregion
            #region //проверка физик или юр лицо
            string namekontrBD_2, familkontrBD;
            CommandBD_zap.CommandText = "SELECT * FROM Контрагент";
            OleDbDataReader dr6 = CommandBD_zap.ExecuteReader();
            while (dr6.Read())
            {
                namekontrBD_2 = dr6.GetString(2);
                if (nameKontr == namekontrBD_2)
                {
                    familkontrBD = dr6.GetString(3);
                    if (familkontrBD != "-") { nalich_famil = true; }
                }
            }
            dr6.Close();
            if (nalich_famil == true) { type_der = "ФИЗ"; }
            else { type_der = "ЮР"; }
            #endregion
            #region //проверка покупка или продажа паев
                if (type_oper == "Покупка паев")
            {
                kol_pay_all = kol_pay_all + kol_pay;
                kol_pay_all = Math.Round(kol_pay_all, 6);
            }
            else if (type_oper == "Продажа паев")
            {
                kol_pay_all = kol_pay_all - kol_pay;
                kol_pay_all = Math.Round(kol_pay_all, 6);
            }
            #endregion

            //Как проверить гражданство?????
            #region //запись строк в БД
            if (kontroz_nalichiy == 1) //значит нет совпадения, нужна новая строка
            {
                string New_paychik = "INSERT INTO [Пайщики_фонда] ([УК], [Фонд], [Счет], [ФИО], [Тип_держателя], [Кол_паев_всего]) VALUES ('" + UK + "', '" + Fond + "', '" + Shet + "', '" + nameKontr + "', '" + type_der + "', '" + kol_pay + "')";
                using (OleDbCommand CommandBDParams = new OleDbCommand(New_paychik, conn_zap))
                { CommandBDParams.ExecuteNonQuery(); }
            }
            if (kontroz_nalichiy == 0) //значит наличие такой фамилии найдено, только изменение кол паев
            {
                string Update_paychik = "UPDATE Пайщики_фонда SET [Кол_паев_всего] = ? WHERE [ФИО] = ? AND [УК] = ? AND [Фонд] = ?";
                using (OleDbCommand CommandBDParams = new OleDbCommand(Update_paychik, conn_zap))
                {
                    CommandBDParams.Parameters.Add("@U1", OleDbType.Double).Value = kol_pay_all;
                    CommandBDParams.Parameters.Add("@U2", OleDbType.Char).Value = nameKontr;
                    CommandBDParams.Parameters.Add("@U3", OleDbType.Char).Value = UK;
                    CommandBDParams.Parameters.Add("@U4", OleDbType.Char).Value = Fond;
                    CommandBDParams.ExecuteNonQuery();
                }
            }
            #endregion
            #region //всегда происходит
            CommandBD_zap.CommandText = "INSERT INTO [Предоставляет_паи] ([ИД_фонд], [ФИО], [Дата], [Кол_паев]) VALUES ('" + ID_fond + "', '" + nameKontr + "', '" + date + "', '" + kol_pay + "')";
            CommandBD_zap.ExecuteNonQuery();
            #endregion
            conn_zap.Close();
            objWorkBook.Close();
            objWorkExcel.Quit();
        }

        /// <summary>
        /// Цикличный метод загрузки в бд
        /// </summary>
        protected internal void Load_doc_cycle()
        {
            var dir = new DirectoryInfo(@".\\Загрузка");                        //папка Загрузка
            string putyiz = new DirectoryInfo(@".\\Загрузка").FullName;                    //путь откуда
            string putyv = new DirectoryInfo(@".\\Ошибки").FullName;                       //путь куда
            string putyarhiv = new DirectoryInfo(@".\\Архив").FullName;      //было везде F:\Учеба\Программа\Архив
            string putyarhiv_from_CB = new DirectoryInfo(@".\\Архив_ошибок_в_отчетах").FullName;


            foreach (FileInfo file in dir.GetFiles("*.xlsx"))                                  //извлекаем все файлы excel
            {
                string nazvfile = file.Name;
                string putfile = file.FullName;
                string nazvaniefilebez = Path.GetFileNameWithoutExtension(file.Name);           //получаем название каждого файла без расширения
                string nazvtxt = nazvaniefilebez + ".txt";                                      //название нового файла ошибок txt
                string novputtxt = Path.Combine(putyv, nazvtxt);                               //новый путь для файла ошибок
                string starputy = Path.Combine(putyiz, nazvfile);                               //старый путь excel файла
                string novputy = Path.Combine(putyv, nazvfile);                                 //новый путь excel файла
                string novputyarhiv = Path.Combine(putyarhiv, nazvfile);
                int kodproverkinazvfile = 0;
                //максимум строк в справочных таблицах (УК, фонд, список....)
                int max_kol_strok_for_proverka = 20;

                #region            //Подключение к БД
                OleDbCommand CommandBD = new OleDbCommand();                                      //команда, через которую все делается
                string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DataBaseMy.accdb";
                OleDbConnection conn = new OleDbConnection(connectionString);                       //новое подключение к БД
                CommandBD.Connection = conn;                                                      //соединение с бд
                conn.Open();
                #endregion
                //Подключение к Excel в каждом случае свое

                #region проверка на согласие
                int nomoshUK_1 = 100, nomoshUK_2 = 100, nomoshF_1 = 100, nomoshF_2 = 100, nomoshRasp = 100,
                    nomoshOsn = 100, nomoshKontr = 100, nomoshfull = 100, nomoshSumm = 100, nomoshVal = 100,
                    nomoshImu = 100, nomoshInoe = 100, nomoshSrok = 100, nomoshDate = 100;

                if (nazvaniefilebez.Contains("согласие"))
                {
                    kodproverkinazvfile = 1;
                    Excel.Application objWorkExcel = new Excel.Application();                   //подключим excel
                    Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(putfile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Excel.Worksheet objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1];      //получим 1 лист

                    #region //УК+лицензия
                    Excel.Range CellUK = objWorkSheet.get_Range("F2");
                    string nameUK = CellUK.Text.ToString();
                    Excel.Range CellLi = objWorkSheet.get_Range("F4");
                    string nameLi = CellLi.Text.ToString();
                    string nameUKBD, nameLiBD, nom_cell_osh_UK = "", nom_cell_osh_lic = "";
                    int nameIDpomoch = 1, nameID, nameIDUK = 0, str_1 = 0;
                    CommandBD.CommandText = "SELECT * FROM Список_Управляющих_компаний";                     //выбор всего из таблицы UK
                    OleDbDataReader dr1 = CommandBD.ExecuteReader();                 //все начало делаться через путь выше
                    while (dr1.Read())
                    {
                        nameUKBD = dr1.GetString(3);                                 //считать 4 столбец  - название УК (начиная с 0)
                        nameLiBD = dr1.GetString(5);                                 //считать 6 столбец - лицензия
                        if (nameLiBD == nameLi)                                     //сравнение лицензии
                        {
                            nomoshUK_2 = 0;
                        }
                        else
                        {
                            nomoshUK_2 = 2;
                            nom_cell_osh_lic = "F4";
                        }
                        if (nameUKBD == nameUK)                                    //сравнение названия УК в БД и excel файле
                        {
                            nomoshUK_1 = 0;
                            nameID = dr1.GetInt32(2);                                //считать ID_UK
                            nameIDpomoch = nameID;
                            nameIDUK = nameID;
                            str_1 = 0;
                            break;
                        }
                        else
                        {
                            str_1++;
                        }
                        if (str_1 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_1 > 0)
                    {
                        nomoshUK_1 = 1;
                        nomoshUK_2 = 0;
                        nom_cell_osh_UK = "F2";
                    }
                    dr1.Close();
                    #endregion
                    #region //Дата
                    string nom_cell_date = "A8";
                    Excel.Range CellDate = objWorkSheet.get_Range("A8");
                    string nameDate = CellDate.Text.ToString();
                    DateTime datasogl = new DateTime();
                    if (nameDate == "")
                    {
                        nomoshDate = 1;
                    }
                    else
                    {
                        nomoshDate = 0;
                        datasogl = DateTime.Parse(nameDate);
                    }
                    #endregion
                    #region //Фонд
                    Excel.Range CellFond = objWorkSheet.get_Range("D13");
                    string nameFond = CellFond.Text.ToString();
                    string nameFBD, nom_cell_osh_fond = "";
                    int nameIDF = 0, str_2 = 0, nameIDUKFond;
                    CommandBD.CommandText = "SELECT * FROM Список_фондов";
                    OleDbDataReader dr2 = CommandBD.ExecuteReader();
                    while (dr2.Read())
                    {
                        nameFBD = dr2.GetString(2);                                 //считать 4 столбец  - название Фонда (начиная с 0)
                        nameIDUKFond = dr2.GetInt32(0);
                        if (nameIDUKFond == nameIDpomoch)
                        {
                            nomoshF_2 = 0;
                        }
                        else
                        {
                            nomoshF_2 = 2;
                            nom_cell_osh_fond = "D13";
                        }
                        if (nameFBD == nameFond)                                    //сравнение названия фонда в БД и excel файле
                        {
                            nomoshF_1 = 0;
                            nameID = dr2.GetInt32(1);                                //считать ID
                            nameIDF = nameID;
                            str_2 = 0;
                            break;
                        }
                        else
                        {
                            str_2++;
                        }
                        if (str_2 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_2 > 0)
                    {
                        nomoshF_1 = 1;
                        nom_cell_osh_fond = "D13";
                    }
                    dr2.Close();
                    #endregion
                    #region //Распоряжение на имущество
                    Excel.Range CellRasp = objWorkSheet.get_Range("A16");
                    string nameRasp = CellRasp.Text.ToString();
                    string nameRaspBD = null, nom_cell_osh_raspor = "";
                    int nameIDRasp = 0, str_3 = 0;
                    CommandBD.CommandText = "SELECT * FROM Справочник_сделок";
                    OleDbDataReader dr3 = CommandBD.ExecuteReader();
                    while (dr3.Read())
                    {
                        nameRaspBD = dr3.GetString(1);                                 //считать столбец  
                        if (nameRaspBD == nameRasp)                                    //сравнение названия в БД и excel файле
                        {
                            nomoshRasp = 0;
                            nameID = dr3.GetInt32(0);                                //считать ID
                            nameIDRasp = nameID;
                            str_3 = 0;
                            break;
                        }
                        else
                        {
                            str_3++;
                        }
                        if (str_3 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_3 > 0)
                    {
                        nomoshRasp = 1;
                        nom_cell_osh_raspor = "A16";
                    }
                    dr3.Close();
                    #endregion
                    #region //Основание
                    Excel.Range CellOsn = objWorkSheet.get_Range("A20");
                    string nameOsn = CellOsn.Text.ToString();
                    string nameOsnBD = null, nom_cell_osh_osnov = "";
                    int nameIDOsn = 0, str_4 = 0;
                    CommandBD.CommandText = "SELECT * FROM Справочник_оснований";
                    OleDbDataReader dr4 = CommandBD.ExecuteReader();
                    while (dr4.Read())
                    {
                        nameOsnBD = dr4.GetString(1);                                 //считать столбец  
                        if (nameOsn == nameOsnBD)                                    //есть ли часть названия в БД из excel файле
                        {
                            nomoshOsn = 0;
                            nameID = dr4.GetInt32(0);                                //считать ID
                            nameIDOsn = nameID;
                            str_4 = 0;
                            break;
                        }
                        else
                        {
                            str_4++;
                        }
                        if (str_4 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_4 > 0)
                    {
                        nomoshOsn = 1;
                        nom_cell_osh_osnov = "A20";
                    }
                    dr4.Close();
                    #endregion
                    #region //контрагент
                    Excel.Range CellKontr = objWorkSheet.get_Range("A24");
                    string nameKontr = CellKontr.Text.ToString();
                    string nameKontrBD = null, nom_cell_osh_kontr = "";
                    int nameIDKontr = 0, str_5 = 0;
                    CommandBD.CommandText = "SELECT * FROM Контрагент";
                    OleDbDataReader dr5 = CommandBD.ExecuteReader();
                    while (dr5.Read())
                    {
                        nameKontrBD = dr5.GetString(2);                                 //считать столбец  
                        if (nameKontr == nameKontrBD)                                    //есть ли часть названия в БД из excel файле
                        {
                            nomoshKontr = 0;
                            nameID = dr5.GetInt32(1);                                //считать ID
                            nameIDKontr = nameID;
                            str_5 = 0;
                            break;
                        }
                        else
                        {
                            str_5++;
                        }
                        if (str_5 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_5 > 0)
                    {
                        nomoshKontr = 1;
                        nom_cell_osh_kontr = "A24";
                    }
                    dr5.Close();
                    #endregion
                    #region //имущество
                    string nom_cell_Imu = "A29";
                    Excel.Range CellImu = objWorkSheet.get_Range("A29");
                    string nameImu = CellImu.Text.ToString();
                    if (nameImu == "")
                    {
                        nomoshImu = 1;
                    }
                    else
                    {
                        nomoshImu = 0;
                    }
                    #endregion
                    #region //иные условия/назначение
                    string nom_cell_Inoe = "A41";
                    Excel.Range CellInoe = objWorkSheet.get_Range("A41");
                    string nameInoe = CellInoe.Text.ToString();
                    if (nameInoe == "")
                    {
                        nomoshInoe = 1;
                    }
                    else
                    {
                        nomoshInoe = 0;
                    }
                    #endregion
                    #region //срок исполнения
                    string nom_cell_srok = "A36";
                    Excel.Range CellSrok = objWorkSheet.get_Range("A36");
                    string nameSrok = CellSrok.Text.ToString();
                    DateTime datasrok = new DateTime();
                    if (nameSrok == "")
                    {
                        nomoshSrok = 1;
                    }
                    else
                    {
                        nomoshSrok = 0;
                        datasrok = DateTime.Parse(nameSrok);
                    }
                    #endregion
                    #region //уникальный код (номер)
                    string cell_nom_nomer = "H10";
                    Excel.Range CellNomer = objWorkSheet.get_Range("H10");
                    string nameNom = CellNomer.Text.ToString();
                    int unikNomer = 0;
                    if (nameNom == "")
                    {
                        nomoshfull = 1;
                    }
                    else
                    {
                        nomoshfull = 0;
                        unikNomer = Convert.ToInt32(nameNom);
                    }
                    #endregion
                    #region //сумма + валюта
                    string nom_cell_sum = "A33";
                    Excel.Range CellSum = objWorkSheet.get_Range("A33");
                    string nameSum = CellSum.Text.ToString();
                    //decimal nameSumm = Decimal.Parse(nameSum);
                    string nom_cell_val = "H33";
                    Excel.Range CellVal = objWorkSheet.get_Range("H33");
                    string nameVal = CellVal.Text.ToString();
                    if (nameSum == "")
                    {
                        nomoshSumm = 1;
                    }
                    else
                    {
                        nomoshSumm = 0;
                        decimal nameSumm = Decimal.Parse(nameSum);
                    }
                    if (nameVal == "")
                    {
                        nomoshVal = 1;
                    }
                    else
                    {
                        nomoshVal = 0;
                    }
                    #endregion
                    //согласовано/несогласовано - без отметок, они проставляются позже

                    /*
                    MessageBox.Show(Convert.ToString(nomoshUK_1), "от УК");
                    MessageBox.Show(Convert.ToString(nomoshF_1), "от фонда");
                    MessageBox.Show(Convert.ToString(nomoshRasp), "от распор");
                    MessageBox.Show(Convert.ToString(nomoshOsn), "от основания");
                    MessageBox.Show(Convert.ToString(nomoshKontr), "от контрагента");
                    */

                    #region //запись в БД
                    if (nomoshUK_1 == 0)
                    {
                        if (nomoshUK_2 == 0)
                        {
                            if (nomoshF_1 == 0)
                            {
                                if (nomoshF_2 == 0)
                                {
                                    if (nomoshRasp == 0)
                                    {
                                        if (nomoshOsn == 0)
                                        {
                                            if ((nomoshSrok == 0) && (nomoshInoe == 0) && (nomoshImu == 0) && (nomoshDate == 0))
                                            {
                                                if (nomoshKontr == 0)
                                                {
                                                    if (nomoshfull == 0)
                                                    {
                                                        if ((nomoshSumm == 0) && (nomoshVal == 0))
                                                        {
                                                            CommandBD.CommandText = "INSERT INTO [Согласия] ([Номер], [Управляющая_компания], [Фонд], [Агент], [Распоряжение], [Дата], [Основание], [Сумма], [Описание_имущества], [Срок_исполнения], [Назначение], [Валюта]) VALUES ('" + unikNomer + "', '" + nameUK + "', '" + nameFond + "', '" + nameKontrBD + "', '" + nameRaspBD + "', '" + datasogl + "', '" + nameOsnBD + "', '" + nameSum + "', '" + nameImu + "', '" + datasrok + "', '" + nameInoe + "', '" + nameVal + "')";
                                                            CommandBD.ExecuteNonQuery();
                                                            conn.Close();
                                                            objWorkBook.Close();
                                                            objWorkExcel.Quit();
                                                            File.Move(starputy, novputyarhiv);  //в архив
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    #endregion

                    #region проверка на ошибки и указание их типа
                    int kol_osh = 0;

                    if (nomoshKontr == 1)
                    {
                        kol_osh++;
                        string text_osh = "Наименование Контрагента не найдено в БД!" + " Ячейка: " + nom_cell_osh_kontr;
                        Zapis_text(novputtxt, text_osh);
                    }
                    if (nomoshOsn == 1)
                    {
                        kol_osh++;
                        string text_osh = "Неопознанный тип основания!" + " Ячейка: " + nom_cell_osh_osnov;
                        Zapis_text(novputtxt, text_osh);
                    }
                    if (nomoshRasp == 1)
                    {
                        kol_osh++;
                        string text_osh = "Неопознанный тип распоряжения!" + " Ячейка: " + nom_cell_osh_raspor;
                        Zapis_text(novputtxt, text_osh);
                    }
                    if (nomoshF_1 == 1)
                    {
                        kol_osh++;
                        string text_osh = "Наименование Фонда не найдено в БД!" + " Ячейка: " + nom_cell_osh_fond;
                        Zapis_text(novputtxt, text_osh);
                    }
                    if (nomoshF_2 == 2)
                    {
                        kol_osh++;
                        string text_osh = "Наименование Фонда не соответствует наименованию УК!" + " Ячейка: " + nom_cell_osh_fond;
                        Zapis_text(novputtxt, text_osh);
                    }
                    if (nomoshUK_1 == 1)
                    {
                        kol_osh++;
                        string text_osh = "Наименование УК не найдено в БД!" + " Ячейка: " + nom_cell_osh_UK;
                        Zapis_text(novputtxt, text_osh);
                    }
                    if (nomoshUK_2 == 2)
                    {
                        kol_osh++;
                        string text_osh = "Номер лицензии не соответствует наименованию УК!" + " Ячейка: " + nom_cell_osh_lic;
                        Zapis_text(novputtxt, text_osh);
                    }
                    if (nomoshfull == 1)
                    {
                        kol_osh++;
                        string text_osh = "Ячейка не заполнена!" + " Ячейка: " + cell_nom_nomer;
                        Zapis_text(novputtxt, text_osh);
                    }
                    if (nomoshSumm == 1)
                    {
                        kol_osh++;
                        string text_osh = "Ячейка не заполнена!" + " Ячейка: " + nom_cell_sum;
                        Zapis_text(novputtxt, text_osh);
                    }
                    if (nomoshVal == 1)
                    {
                        kol_osh++;
                        string text_osh = "Ячейка не заполнена!" + " Ячейка: " + nom_cell_val;
                        Zapis_text(novputtxt, text_osh);
                    }
                    if (nomoshImu == 1)
                    {
                        kol_osh++;
                        string text_osh = "Ячейка не заполнена!" + " Ячейка: " + nom_cell_Imu;
                        Zapis_text(novputtxt, text_osh);
                    }
                    if (nomoshInoe == 1)
                    {
                        kol_osh++;
                        string text_osh = "Ячейка не заполнена!" + " Ячейка: " + nom_cell_Inoe;
                        Zapis_text(novputtxt, text_osh);
                    }
                    if (nomoshSrok == 1)
                    {
                        kol_osh++;
                        string text_osh = "Ячейка не заполнена!" + " Ячейка: " + nom_cell_srok;
                        Zapis_text(novputtxt, text_osh);
                    }
                    if (nomoshDate == 1)
                    {
                        kol_osh++;
                        string text_osh = "Ячейка не заполнена!" + " Ячейка: " + nom_cell_date;
                        Zapis_text(novputtxt, text_osh);
                    }

                    if (kol_osh > 0)
                    {
                        objWorkBook.Close();
                        objWorkExcel.Quit();
                        File.Move(starputy, novputy);                                                  // в ошибки
                    }

                    #endregion
                    objWorkExcel.Quit();

                }
                #endregion

                #region проверка на плат поручение
                int nomeroshUK = 1, nomeroshUK_shet = 1, nomeroshF_1 = 1, nomeroshF_2 = 1, nomeroshBplat_1 = 1, nomeroshBplat_2 = 1,
                    nomeroshBplat_3 = 1, nomeroshBpol_1 = 1, nomeroshBpol_2 = 100, nomeroshBpol_3 = 100, nomeroshKontr_1 = 1,
                    nomeroshKontr_2 = 1, nomeroshRasp = 1, nomeroshVal = 1, nomeroshDate = 1, nomeroshInoe = 1,
                    nomeroshNomer = 1, nomeroshVid = 1, nomeroshSum = 1, nomeroshSumprop = 1;

                if (nazvaniefilebez.Contains("плат_пор"))
                {
                    kodproverkinazvfile = 2;
                    Excel.Application objWorkExcel = new Excel.Application();                   //подключим excel
                    Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(putfile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Excel.Worksheet objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1];      //получим 1 лист

                    #region //УК + счет
                    Excel.Range CellUK = objWorkSheet.get_Range("A12");
                    string nameUK = CellUK.Text.ToString();
                    Excel.Range CellUKShet = objWorkSheet.get_Range("U13");
                    string nameUKShet = CellUKShet.Text.ToString();

                    string nameUKBD, nameUKshetBD, nomer_cell_osh_UK = "", nomer_cell_UK_shet = "";
                    int nameIDpomoch = 1, nameID;
                    int nameIDUK = 0, str_6 = 0;
                    CommandBD.CommandText = "SELECT * FROM Список_Управляющих_компаний";                     //выбор всего из таблицы UK
                    OleDbDataReader dr1 = CommandBD.ExecuteReader();                 //все начало делаться через путь выше
                    while (dr1.Read())
                    {
                        nameUKBD = dr1.GetString(3);                                 //считать 4 столбец  - название УК (начиная с 0)
                        nameUKshetBD = dr1.GetString(10);
                        if (nameUKShet == nameUKshetBD)
                        {
                            nomeroshUK_shet = 0;
                        }
                        else
                        {
                            nomeroshUK_shet = 2;
                            nomer_cell_UK_shet = "U13";
                        }
                        if (nameUKBD == nameUK)                                    //сравнение названия УК в БД и excel файле
                        {
                            nomeroshUK = 0;
                            nameID = dr1.GetInt32(2);                                //считать ID_UK
                            nameIDpomoch = nameID;
                            nameIDUK = nameID;
                            str_6 = 0;
                            break;
                        }
                        else
                        {
                            str_6++;
                        }
                        if (str_6 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_6 > 0)
                    {
                        nomeroshUK = 1;
                        nomer_cell_osh_UK = "A12";
                    }
                    dr1.Close();
                    #endregion
                    #region //Дата
                    string nomer_cell_date = "O6";
                    Excel.Range CellDate = objWorkSheet.get_Range("O6");
                    string nameDate = CellDate.Text.ToString();
                    DateTime dataplat = new DateTime();
                    if (nameDate == "")
                    {
                        nomeroshDate = 1;
                    }
                    else
                    {
                        nomeroshDate = 0;
                        dataplat = DateTime.Parse(nameDate);
                    }
                    #endregion
                    #region //Фонд
                    Excel.Range CellFond = objWorkSheet.get_Range("A13");
                    string nameFond = CellFond.Text.ToString();

                    string nameFBD, nomer_cell_osh_fond = "";
                    int nameIDF = 0, str_7 = 0;
                    int nameIDUKFond;
                    CommandBD.CommandText = "SELECT * FROM Список_фондов";
                    OleDbDataReader dr2 = CommandBD.ExecuteReader();
                    while (dr2.Read())
                    {
                        nameFBD = dr2.GetString(2);                                 //считать 4 столбец  - название Фонда (начиная с 0)
                        nameIDUKFond = dr2.GetInt32(0);
                        if (nameIDUKFond == nameIDpomoch)
                        {
                            nomeroshF_2 = 0;
                        }
                        else
                        {
                            nomeroshF_2 = 2;
                            nomer_cell_osh_fond = "A13";
                        }
                        if (nameFBD == nameFond)                                    //сравнение названия фонда в БД и excel файле
                        {
                            nomeroshF_1 = 0;
                            nameID = dr2.GetInt32(1);                                //считать ID
                            nameIDF = nameID;
                            str_7 = 0;
                            break;
                        }
                        else
                        {
                            str_7++;
                        }
                        if (str_7 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_7 > 0)
                    {
                        nomeroshF_1 = 1;
                        nomeroshF_2 = 0;
                        nomer_cell_osh_fond = "A13";
                    }
                    dr2.Close();
                    #endregion
                    #region //Распоряжение на имущество
                    Excel.Range CellRasp = objWorkSheet.get_Range("A33");
                    string nameRasp = CellRasp.Text.ToString();

                    string nameRaspBD = null, nomer_cell_osh_raspor = "";
                    int nameIDRasp = 0, str_8 = 0;
                    CommandBD.CommandText = "SELECT * FROM Справочник_сделок";
                    OleDbDataReader dr3 = CommandBD.ExecuteReader();
                    while (dr3.Read())
                    {
                        nameRaspBD = dr3.GetString(1);                                 //считать столбец  
                        if (nameRaspBD == nameRasp)                                    //сравнение названия в БД и excel файле
                        {
                            nomeroshRasp = 0;
                            nameID = dr3.GetInt32(0);                                //считать ID
                            nameIDRasp = nameID;
                            str_8 = 0;
                            break;
                        }
                        else
                        {
                            str_8++;
                        }
                        if (str_8 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_8 > 0)
                    {
                        nomeroshRasp = 1;
                        nomer_cell_osh_raspor = "A33";
                    }
                    dr3.Close();
                    #endregion
                    #region //контрагент
                    Excel.Range CellKontr = objWorkSheet.get_Range("A25");
                    string nameKontr = CellKontr.Text.ToString();
                    Excel.Range CellKontrShet = objWorkSheet.get_Range("U24");
                    string nameKontrShet = CellKontrShet.Text.ToString();

                    string nameKontrBD, nameKontrShetBD, nomer_cell_osh_kontr = "", nomer_cell_osh_kontr_chet = "";
                    int nameIDKontr = 0, str_9 = 0;
                    CommandBD.CommandText = "SELECT * FROM Контрагент";
                    OleDbDataReader dr5 = CommandBD.ExecuteReader();
                    while (dr5.Read())
                    {
                        nameKontrBD = dr5.GetString(2);                                //считать столбец  
                        nameKontrShetBD = dr5.GetString(6);
                        if (nameKontrShet == nameKontrShetBD)
                        {
                            nomeroshKontr_2 = 0;
                        }
                        else
                        {   
                            nomeroshKontr_2 = 2;
                            nomer_cell_osh_kontr_chet = "U24";
                        }
                        if (nameKontr == nameKontrBD)                                    //есть ли часть названия в БД из excel файле
                        {
                            nomeroshKontr_1 = 0;
                            nameID = dr5.GetInt32(1);                                //считать ID
                            nameIDKontr = nameID;
                            str_9 = 0;
                            break;
                        }
                        else
                        {
                            str_9++;
                        }
                        if (str_9 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_9 > 0)
                    {
                        nomeroshKontr_1 = 1;
                        nomeroshKontr_2 = 0;
                        nomer_cell_osh_kontr = "A25";
                    }
                    dr5.Close();
                    #endregion
                    #region //Назначение - сам текст!
                    string nomer_cell_Inoe = "G33";
                    Excel.Range CellInoe = objWorkSheet.get_Range("G33");
                    string nameInoe = CellInoe.Text.ToString();
                    if (nameInoe == "")
                    {
                        nomeroshInoe = 1;
                    }
                    else
                    {
                        nomeroshInoe = 0;
                    }
                    #endregion
                    #region //валюта
                    Excel.Range CellVal = objWorkSheet.get_Range("W2");
                    string nameVal = CellVal.Text.ToString();
                    string nameValBD, nomer_cell_osh_val = "";
                    int nameIDVal = 0, str_10 = 0;
                    CommandBD.CommandText = "SELECT * FROM Справочник_валют";
                    OleDbDataReader dr6 = CommandBD.ExecuteReader();
                    while (dr6.Read())
                    {
                        nameValBD = dr6.GetString(1);
                        if (nameVal.Contains(nameValBD))
                        {
                            nomeroshVal = 0;
                            nameID = dr6.GetInt32(0);
                            nameIDVal = nameID;
                            str_10 = 0;
                            break;
                        }
                        else
                        {
                            str_10++;
                        }
                        if (str_10 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_10 > 0)
                    {
                        nomeroshVal = 1;
                        nomer_cell_osh_val = "W2";
                    }
                    dr6.Close();
                    #endregion
                    #region //номер плат поручения
                    string nomer_cell_nomer = "J6";
                    Excel.Range CellNom = objWorkSheet.get_Range("J6");
                    string nameNom = CellNom.Text.ToString();
                    int nameNomerPP = new int();
                    if (nameNom == "")
                    {
                        nomeroshNomer = 1;
                    }
                    else
                    {
                        nomeroshNomer = 0;
                        nameNomerPP = Convert.ToInt32(nameNom);
                    }
                    #endregion
                    #region //вид платежа (дебет/кредит)
                    string nomer_cell_vid = "W6";
                    Excel.Range CellVid = objWorkSheet.get_Range("W6");
                    string nameVid = CellVid.Text.ToString();
                    if (nameVid == "")
                    {
                        nomeroshVid = 1;
                    }
                    else
                    {
                        nomeroshVid = 0;
                    }
                    #endregion
                    #region //сумма
                    string nomer_cell_sum = "U11";
                    Excel.Range CellSum = objWorkSheet.get_Range("U11");
                    string nameSum = CellSum.Text.ToString();
                    decimal nameSumm = new decimal();
                    if (nameSum == "")
                    {
                        nomeroshSum = 1;
                    }
                    else
                    {
                        nomeroshSum = 0;
                        nameSumm = Decimal.Parse(nameSum);
                    }
                    string nomer_cell_sum_prop = "C10";
                    Excel.Range CellSumprop = objWorkSheet.get_Range("C10");
                    string nameSumprop = CellSumprop.Text.ToString();
                    if (nameSumprop == "")
                    {
                        nomeroshSumprop = 1;
                    }
                    else
                    {
                        nomeroshSumprop = 0;
                    }
                    #endregion
                    #region //банк плательщика
                    Excel.Range CellBankPl = objWorkSheet.get_Range("A16");
                    string nameBankPl = CellBankPl.Text.ToString();
                    Excel.Range CellBankPlBik = objWorkSheet.get_Range("U16");
                    string nameBankPlBik = CellBankPlBik.Text.ToString();
                    Excel.Range CellBankPlKor = objWorkSheet.get_Range("U17");
                    string nameBankPlKor = CellBankPlKor.Text.ToString();

                    string nameBankPlBD, nameBankPlBikBD, nameBankPlKorBD, nomer_cell_osh_Bplat = "", nomer_cell_osh_Bplat_bik = "", nomer_cell_osh_Bplat_kor = "";
                    int nameIDBankPl = 0, str_11 = 0;
                    CommandBD.CommandText = "SELECT * FROM Список_банков";
                    OleDbDataReader dr7 = CommandBD.ExecuteReader();
                    while (dr7.Read())
                    {
                        nameBankPlBD = dr7.GetString(1);                                 //считать столбец 
                        nameBankPlBikBD = dr7.GetString(2);
                        nameBankPlKorBD = dr7.GetString(3);
                        if (nameBankPlKor == nameBankPlKorBD)
                        {
                            nomeroshBplat_3 = 0;
                        }
                        else
                        {
                            nomeroshBplat_3 = 3;
                            nomer_cell_osh_Bplat_kor = "U17";
                        }
                        if (nameBankPlBik == nameBankPlBikBD)
                        {
                            nomeroshBplat_2 = 0;
                        }
                        else
                        {
                            nomeroshBplat_2 = 2;
                            nomer_cell_osh_Bplat_bik = "U16";
                        }
                        if (nameBankPl == nameBankPlBD)                                    //есть ли часть названия в БД из excel файле
                        {
                            nomeroshBplat_1 = 0;
                            nameID = dr7.GetInt32(0);                                //считать ID
                            nameIDBankPl = nameID;
                            str_11 = 0;
                            break;
                        }
                        else
                        {
                            str_11++;
                        }
                        if (str_11 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_11 > 0)
                    {
                        nomeroshBplat_1 = 1;
                        nomeroshBplat_2 = 0;
                        nomeroshBplat_3 = 0;
                        nomer_cell_osh_Bplat = "A16";
                    }
                    dr7.Close();
                    #endregion
                    #region //банк получателя
                    Excel.Range CellBankPol = objWorkSheet.get_Range("A20");
                    string nameBankPol = CellBankPol.Text.ToString();
                    Excel.Range CellBankPolBik = objWorkSheet.get_Range("U20");
                    string nameBankPolBik = CellBankPolBik.Text.ToString();
                    Excel.Range CellBankPolKor = objWorkSheet.get_Range("U21");
                    string nameBankPolKor = CellBankPolKor.Text.ToString();

                    string nameBankPolBD, nameBankPolBikBD, nameBankPolKorBD, nomer_cell_osh_Bpol = "", nomer_cell_osh_Bpol_bik = "", nomer_cell_osh_Bpol_kor = "";
                    int nameIDBankPol = 0, str_12 = 0;
                    CommandBD.CommandText = "SELECT * FROM Список_банков";
                    OleDbDataReader dr8 = CommandBD.ExecuteReader();
                    while (dr8.Read())
                    {
                        nameBankPolBD = dr8.GetString(1);                                 //считать столбец 
                        nameBankPolBikBD = dr8.GetString(2);
                        nameBankPolKorBD = dr8.GetString(3);
                        if (nameBankPolKor == nameBankPolKorBD)
                        {
                            nomeroshBpol_3 = 0;
                        }
                        else
                        {
                            nomeroshBpol_3 = 3;
                            nomer_cell_osh_Bpol_kor = "U21";
                        }
                        if (nameBankPolBik == nameBankPolBikBD)
                        {
                            nomeroshBpol_2 = 0;
                        }
                        else
                        {
                            nomeroshBpol_2 = 2;
                            nomer_cell_osh_Bpol_bik = "U20";
                        }
                        if (nameBankPol == nameBankPolBD)                                    //есть ли часть названия в БД из excel файле
                        {
                            nomeroshBpol_1 = 0;
                            nameID = dr8.GetInt32(0);                                //считать ID
                            nameIDBankPol = nameID;
                            str_12 = 0;
                            break;
                        }
                        else
                        {
                            str_12++;
                        }
                        if (str_12 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_12 > 0)
                    {
                        nomeroshBpol_1 = 1;
                        nomeroshBpol_2 = 0;
                        nomeroshBpol_3 = 0;
                        nomer_cell_osh_Bpol = "A20";
                    }
                    dr8.Close();
                    #endregion
                    #region //количество паев
                    Excel.Range CellColP = objWorkSheet.get_Range("G31");
                    string nameCP = CellColP.Text.ToString();
                    decimal nameColP = new decimal();
                    if (nameCP == "")
                    {
                        //MessageBox.Show("Пусто");
                    }
                    else
                    {
                        nameColP = Convert.ToDecimal(nameCP);
                        //nameColP = Decimal.Parse(nameCP);
                    }
                    #endregion

                    /*
                    MessageBox.Show(Convert.ToString(nomeroshUK), "от УК");
                    MessageBox.Show(Convert.ToString(nomeroshF_1), "от фонда");
                    MessageBox.Show(Convert.ToString(nomeroshRasp), "от распор");
                    MessageBox.Show(Convert.ToString(nomeroshVal), "от валюты");
                    MessageBox.Show(Convert.ToString(nomeroshKontr_1), "от контрагента");
                    MessageBox.Show(Convert.ToString(nomeroshBplat_1), "от банка плат");
                    MessageBox.Show(Convert.ToString(nomeroshBpol_1), "от банка пол");
                    */

                    #region //запись в БД без паев!
                    if (nameCP == "")
                    {
                        if ((nomeroshUK == 0) && (nomeroshUK_shet == 0))
                        {
                            if ((nomeroshF_1 == 0) && (nomeroshF_2 == 0))
                            {
                                if ((nomeroshDate == 0) && (nomeroshInoe == 0) && (nomeroshNomer == 0) && (nomeroshVid == 0) && (nomeroshSumprop == 0) && (nomeroshSum == 0))
                                {
                                    if (nomeroshRasp == 0)
                                    {
                                        if (nomeroshVal == 0)
                                        {
                                            if ((nomeroshKontr_1 == 0) && (nomeroshKontr_2 == 0))
                                            {
                                                if ((nomeroshBplat_1 == 0) && (nomeroshBplat_2 == 0) && (nomeroshBplat_3 == 0))
                                                {
                                                    if ((nomeroshBpol_1 == 0) && (nomeroshBpol_2 == 0) && (nomeroshBpol_3 == 0))
                                                    {
                                                        CommandBD.CommandText = "INSERT INTO [Платежные_поручения] ([Номер_пп], [Управляющая_компания], [Фонд], [Агент], [Распоряжение], [Дата], [Валюта], [Сумма], [Вид_платежа], [Кол_паев], [Назначение], [Банк_плат], [Банк_получ]) VALUES ('" + nameNomerPP + "', '" + nameUK + "', '" + nameFond + "', '" + nameKontr + "', '" + nameRaspBD + "', '" + dataplat + "', '" + nameVal + "', '" + nameSumm + "', '" + nameVid + "', '" + nameColP + "', '" + nameInoe + "', '" + nameBankPl + "', '" + nameBankPol + "')";
                                                        CommandBD.ExecuteNonQuery();
                                                        conn.Close();
                                                        objWorkBook.Close();
                                                        objWorkExcel.Quit();
                                                        File.Move(starputy, novputyarhiv);      //в архив
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                    #endregion
                    #region Запись в БД если есть паи
                    else
                    {
                        #region
                        if ((nomeroshUK == 0) && (nomeroshUK_shet == 0))
                        {
                            if ((nomeroshF_1 == 0) && (nomeroshF_2 == 0))
                            {
                                if ((nomeroshDate == 0) && (nomeroshInoe == 0) && (nomeroshNomer == 0) && (nomeroshVid == 0) && (nomeroshSumprop == 0) && (nomeroshSum == 0))
                                {
                                    if (nomeroshRasp == 0)
                                    {
                                        if (nomeroshVal == 0)
                                        {
                                            if ((nomeroshKontr_1 == 0) && (nomeroshKontr_2 == 0))
                                            {
                                                if ((nomeroshBplat_1 == 0) && (nomeroshBplat_2 == 0) && (nomeroshBplat_3 == 0))
                                                {
                                                    if ((nomeroshBpol_1 == 0) && (nomeroshBpol_2 == 0) && (nomeroshBpol_3 == 0))
                                                    {
                                                        CommandBD.CommandText = "INSERT INTO [Платежные_поручения] ([Номер_пп], [Управляющая_компания], [Фонд], [Агент], [Распоряжение], [Дата], [Валюта], [Сумма], [Вид_платежа], [Кол_паев], [Назначение], [Банк_плат], [Банк_получ]) VALUES ('" + nameNomerPP + "', '" + nameUK + "', '" + nameFond + "', '" + nameKontr + "', '" + nameRaspBD + "', '" + dataplat + "', '" + nameVal + "', '" + nameSumm + "', '" + nameVid + "', '" + nameColP + "', '" + nameInoe + "', '" + nameBankPl + "', '" + nameBankPol + "')";
                                                        CommandBD.ExecuteNonQuery();
                                                        conn.Close();
                                                        objWorkBook.Close();
                                                        objWorkExcel.Quit();
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                        #endregion
                        Zagruzka_paychika(putfile, max_kol_strok_for_proverka, nameIDF, nameFond, nameUK, nameKontrShet, dataplat, nameColP, nameRaspBD);
                        File.Move(starputy, novputyarhiv);      //в архив
                    }
                    #endregion

                    #region проверка на ошибки и указание их типа
                    int kol_osh = 0;

                    if (nomeroshBpol_3 == 3)
                    {
                        kol_osh++;
                        StreamWriter nevtxt7 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt7.WriteLine("Корреспондентский счет получателя не верен!" + " Ячейка: " + nomer_cell_osh_Bpol_kor);
                        nevtxt7.Close();
                    }
                    if (nomeroshBpol_2 == 2)
                    {
                        kol_osh++;
                        StreamWriter nevtxt6 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt6.WriteLine("БИК получателя не верен!" + " Ячейка: " + nomer_cell_osh_Bpol_bik);
                        nevtxt6.Close();
                    }
                    if (nomeroshBpol_1 == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt5 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt5.WriteLine("Наименование банка получателя не найдено в БД!" + " Ячейка: " + nomer_cell_osh_Bpol);
                        nevtxt5.Close();
                    }
                    if (nomeroshBplat_3 == 3)
                    {
                        kol_osh++;
                        StreamWriter nevtxt7 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt7.WriteLine("Корреспондентский счет плательщика не верен!" + " Ячейка: " + nomer_cell_osh_Bplat_kor);
                        nevtxt7.Close();
                    }
                    if (nomeroshBplat_2 == 2)
                    {
                        kol_osh++;
                        StreamWriter nevtxt6 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt6.WriteLine("БИК плательщика не верен!" + " Ячейка: " + nomer_cell_osh_Bplat_bik);
                        nevtxt6.Close();
                    }
                    if (nomeroshBplat_1 == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt5 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt5.WriteLine("Наименование банка плательщика не найдено в БД!" + " Ячейка: " + nomer_cell_osh_Bplat);
                        nevtxt5.Close();
                    }
                    if (nomeroshKontr_2 == 2)
                    {
                        kol_osh++;
                        StreamWriter nevtxt5 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt5.WriteLine("Счет контрагента не найден в БД!" + " Ячейка: " + nomer_cell_osh_kontr_chet);
                        nevtxt5.Close();
                    }
                    if (nomeroshKontr_1 == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt5 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt5.WriteLine("Наименование контрагента не найдено в БД!" + " Ячейка: " + nomer_cell_osh_kontr);
                        nevtxt5.Close();
                    }
                    if (nomeroshVal == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt5 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt5.WriteLine("Наименование валюты не найдено в БД!" + " Ячейка: " + nomer_cell_osh_val);
                        nevtxt5.Close();
                    }
                    if (nomeroshRasp == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt5 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt5.WriteLine("Наименование распоряжения не найдено в БД!" + " Ячейка: " + nomer_cell_osh_raspor);
                        nevtxt5.Close();
                    }
                    if (nomeroshF_2 == 2)
                    {
                        kol_osh++;
                        StreamWriter nevtxt5 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt5.WriteLine("Фонд не соответствует УК!" + " Ячейка: " + nomer_cell_osh_fond);
                        nevtxt5.Close();
                    }
                    if (nomeroshF_1 == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt5 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt5.WriteLine("Наименование фонда не найдено в БД!" + " Ячейка: " + nomer_cell_osh_fond);
                        nevtxt5.Close();
                    }
                    if (nomeroshUK == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt5 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt5.WriteLine("Наименование УК не найдено в БД!" + " Ячейка: " + nomer_cell_osh_UK);
                        nevtxt5.Close();
                    }
                    if (nomeroshUK_shet == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt5 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt5.WriteLine("Расчетный счет УК не найден в БД!" + " Ячейка: " + nomer_cell_UK_shet);
                        nevtxt5.Close();
                    }
                    if (nomeroshVid == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt5 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt5.WriteLine("Не указан вид платежа!" + " Ячейка: " + nomer_cell_vid);
                        nevtxt5.Close();
                    }
                    if (nomeroshSumprop == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt5 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt5.WriteLine("Не указана сумма прописью!" + " Ячейка: " + nomer_cell_sum_prop);
                        nevtxt5.Close();
                    }
                    if (nomeroshSum == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt5 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt5.WriteLine("Не указана сумма!" + " Ячейка: " + nomer_cell_sum);
                        nevtxt5.Close();
                    }
                    if (nomeroshNomer == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt5 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt5.WriteLine("Не указан номер платежного поручения!" + " Ячейка: " + nomer_cell_nomer);
                        nevtxt5.Close();
                    }
                    if (nomeroshInoe == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt5 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt5.WriteLine("Не указано назначение платежа!" + " Ячейка: " + nomer_cell_Inoe);
                        nevtxt5.Close();
                    }
                    if (kol_osh > 0)
                    {
                        objWorkBook.Close();
                        objWorkExcel.Quit();
                        File.Move(starputy, novputy);                                                  // в ошибки
                    }

                    #endregion

                    objWorkExcel.Quit();
                    conn.Close();
                }
                #endregion

                #region проверка на отметки ЦБ
                int osh_name_otch = 1, osh_date_otch = 1, osh_UK = 1, osh_UK_lic = 1, osh_city = 1, osh_lic_fond = 1, otmet_osh_uved = 1;
                if (nazvaniefilebez.Contains("отметк"))
                {
                    kodproverkinazvfile = 3;
                    Excel.Application objWorkExcel = new Excel.Application();                   //подключим excel
                    Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(putfile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Excel.Worksheet objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1];      //получим 1 лист

                    #region //УК+лицензия
                    Excel.Range CellUK = objWorkSheet.get_Range("F17");
                    string nameUK = CellUK.Text.ToString();
                    Excel.Range CellLi = objWorkSheet.get_Range("F19");
                    string nameLi = CellLi.Text.ToString();

                    string nameUKBD, nameLiBD, nom_cell_osh_UK = "", nom_cell_osh_lic = "";
                    int nameIDpomoch = 1, nameID;
                    int nameIDUK = 0;
                    int n = 1, str_1 = 0;
                    CommandBD.CommandText = "SELECT * FROM Список_Управляющих_компаний";                     //выбор всего из таблицы UK
                    OleDbDataReader dr1 = CommandBD.ExecuteReader();                 //все начало делаться через путь выше
                    while (dr1.Read())
                    {
                        nameUKBD = dr1.GetString(3);                                 //считать 4 столбец  - название УК (начиная с 0)
                        nameLiBD = dr1.GetString(5);                                 //считать 6 столбец - лицензия
                        if (nameLiBD == nameLi)                                     //сравнение лицензии
                        {
                            osh_UK_lic = 0;
                        }
                        else
                        {
                            osh_UK_lic = 2;
                            nom_cell_osh_lic = "F19";
                        }
                        if (nameUKBD == nameUK)                                    //сравнение названия УК в БД и excel файле
                        {
                            osh_UK = 0;
                            nameID = dr1.GetInt32(2);                                //считать ID_UK
                            nameIDpomoch = nameID;
                            nameIDUK = nameID;
                            str_1 = 0;
                            break;
                        }
                        else
                        {
                            str_1++;
                        }
                        if (str_1 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_1 > 0)
                    {
                        osh_UK = 1;
                        osh_UK_lic = 0;
                        nom_cell_osh_UK = "F17";
                    }
                    dr1.Close();
                    #endregion
                    #region //Дата
                    string nom_cell_date = "A8";
                    Excel.Range CellDate_otm = objWorkSheet.get_Range("A8");
                    string nameDate_otm = CellDate_otm.Text.ToString();
                    DateTime data_otmet = new DateTime();
                    if (nameDate_otm == "")
                    {
                        nomoshDate = 1;
                    }
                    else
                    {
                        nomoshDate = 0;
                        data_otmet = DateTime.Parse(nameDate_otm);
                    }
                    #endregion
                    #region //Дата отчета
                    string nom_cell_date_otch = "A26";
                    Excel.Range CellDate_otch = objWorkSheet.get_Range("A26");
                    string nameDate_otch = CellDate_otch.Text.ToString();
                    DateTime data_otch = new DateTime();
                    if (nameDate_otch == "")
                    {
                        osh_date_otch = 1;
                    }
                    else
                    {
                        osh_date_otch = 0;
                        data_otch = DateTime.Parse(nameDate_otch);
                    }
                    #endregion
                    #region //Фонд - без лицензии
                    Excel.Range CellFond = objWorkSheet.get_Range("F21");
                    string nameFond = CellFond.Text.ToString();

                    string nameFBD, nom_cell_osh_fond = "";
                    int nameIDF = 0, str_2 = 0;
                    int nameIDUKFond;
                    CommandBD.CommandText = "SELECT * FROM Список_фондов";
                    OleDbDataReader dr2 = CommandBD.ExecuteReader();
                    while (dr2.Read())
                    {
                        nameFBD = dr2.GetString(2);                                 //считать 4 столбец  - название Фонда (начиная с 0)
                        nameIDUKFond = dr2.GetInt32(0);
                        if (nameIDUKFond == nameIDpomoch)
                        {
                            nomoshF_2 = 0;
                        }
                        else
                        {
                            nomoshF_2 = 2;
                            nom_cell_osh_fond = "F21";
                        }
                        if (nameFBD == nameFond)                                    //сравнение названия фонда в БД и excel файле
                        {
                            nomoshF_1 = 0;
                            nameID = dr2.GetInt32(1);                                //считать ID
                            nameIDF = nameID;
                            str_2 = 0;
                            break;
                        }
                        else
                        {
                            str_2++;
                        }
                        if (str_2 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_2 > 0)
                    {
                        nomoshF_1 = 1;
                        nom_cell_osh_fond = "F21";
                    }
                    dr2.Close();
                    #endregion
                    #region //название  отчета
                    string cell_osh_name_otch = "A30";
                    Excel.Range Cell_nameot_osh = objWorkSheet.get_Range("A30");
                    string nameot_osh = Cell_nameot_osh.Text.ToString();
                    osh_name_otch = 0;
                    if (nameot_osh == "")
                    {
                        osh_name_otch = 1;
                    }
                    #endregion
                    #region //текст отказа
                    Excel.Range Cell_text_otk = objWorkSheet.get_Range("A33");
                    string text_otk = Cell_text_otk.Text.ToString();
                    bool otm_otkaz = false;
                    bool otm_prin = false;
                    if (text_otk == "")
                    {
                        otm_prin = true;
                    }
                    else
                    {
                        otm_otkaz = true;
                    }
                    #endregion
                    #region //город + заполнение лицензии (не проверка по БД!)
                    string nom_cell_city = "F18", nom_cell_lic_fond = "F22";
                    Excel.Range CellCity = objWorkSheet.get_Range("F18");
                    string nameCity = CellCity.Text.ToString();
                    Excel.Range Cell_lic_fond = objWorkSheet.get_Range("F22");
                    string name_lic_fond = Cell_lic_fond.Text.ToString();
                    osh_city = 0;
                    osh_lic_fond = 0;
                    if (nameCity == "")
                    {
                        osh_city = 1;
                    }
                    if (name_lic_fond == "")
                    {
                        osh_lic_fond = 1;
                    }
                    #endregion
                    #region //номер уведомления
                    string nom_cell_uved = "H8";
                    Excel.Range Cell_uved = objWorkSheet.get_Range("H8");
                    string nomer_uved = Cell_uved.Text.ToString();
                    otmet_osh_uved = 0;
                    if (nomer_uved == "")
                    {
                        otmet_osh_uved = 1;
                    }
                    #endregion

                    #region //запись в БД выписка
                    if (nameot_osh.Contains("Выписка"))
                    {
                        if ((osh_UK == 0) && (osh_UK_lic == 0))
                        {
                            if ((nomoshF_1 == 0) && (nomoshF_2 == 0))
                            {
                                if ((osh_name_otch == 0) && (nomoshDate == 0) && (osh_date_otch == 0) && (osh_city == 0) && (osh_lic_fond == 0) && (otmet_osh_uved == 0))
                                {
                                    if (otm_prin == true)
                                    {
                                        string Update_Vip = "UPDATE Выписка_день SET [Отметка_о_принятии_ЦБ]= ? , [Дата]=?, [Дата_принятия_ЦБ] = ? WHERE [Название_выписки] = ?";
                                        using (OleDbCommand CommandBDParams = new OleDbCommand(Update_Vip, conn))
                                        {
                                            CommandBDParams.Parameters.Add("@M1", OleDbType.Boolean).Value = true;
                                            CommandBDParams.Parameters.Add("@M2", OleDbType.Date).Value = data_otmet;
                                            CommandBDParams.Parameters.Add("@M3", OleDbType.Date).Value = data_otmet;
                                            CommandBDParams.Parameters.Add("@M4", OleDbType.Char).Value = nameot_osh;
                                            CommandBDParams.ExecuteNonQuery();
                                        }
                                        conn.Close();
                                        objWorkBook.Close();
                                        objWorkExcel.Quit();
                                        File.Move(starputy, novputyarhiv);  //в архив
                                    }
                                    if (otm_otkaz == true)
                                    {
                                        string Update_Vip = "UPDATE Выписка_день SET [Отметка_отказа]= ? , [Дата]=?, [Текст_отказа] = ? WHERE [Название_выписки] = ?";
                                        using (OleDbCommand CommandBDParams = new OleDbCommand(Update_Vip, conn))
                                        {
                                            CommandBDParams.Parameters.Add("@N1", OleDbType.Boolean).Value = true;
                                            CommandBDParams.Parameters.Add("@N2", OleDbType.Date).Value = data_otmet;
                                            CommandBDParams.Parameters.Add("@N3", OleDbType.Char).Value = text_otk;
                                            CommandBDParams.Parameters.Add("@N4", OleDbType.Char).Value = nameot_osh;
                                            CommandBDParams.ExecuteNonQuery();
                                        }
                                        conn.Close();
                                        objWorkBook.Close();
                                        objWorkExcel.Quit();
                                        File.Move(starputy, novputyarhiv);  //в архив
                                    }
                                }
                            }
                        }
                    }
                    #endregion

                    #region //запись в БД сча
                    if (nameot_osh.Contains("СЧА"))
                    {
                        /*if (nomoshUK_1 == 0)
                        {
                            if (nomoshUK_2 == 0)
                            {
                                if (nomoshF_1 == 0)
                                {
                                    if (nomoshF_2 == 0)
                                    {
                                        CommandBD.CommandText = "INSERT INTO [Ошибки_ЦБ] ([Дата], [Номер_уведомления], [Наименование_УК], [Наименование_фонда], [Дата_отчета], [Название_отчета]) VALUES ('" + data_otmet + "', '" + nomer_uved + "', '" + nameUK + "', '" + nameFond + "', '" + data_otch + "', '" + nameot_with_osh + "')";
                                        CommandBD.ExecuteNonQuery();
                                        conn.Close();
                                        objWorkBook.Close();
                                        objWorkExcel.Quit();
                                        File.Move(starputy, novputyarhiv);  //в архив
                                    }
                                }
                            }
                        }*/
                    }
                    #endregion

                    #region проверка на ошибки и указание их типа
                    int kol_osh = 0;

                    if (nomoshF_1 == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt3 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt3.WriteLine("Наименование Фонда не найдено в БД!" + " Ячейка: " + nom_cell_osh_fond);
                        nevtxt3.Close();
                    }
                    if (nomoshF_2 == 2)
                    {
                        kol_osh++;
                        StreamWriter nevtxt4 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt4.WriteLine("Наименование Фонда не соответствует наименованию УК!" + " Ячейка: " + nom_cell_osh_fond);
                        nevtxt4.Close();
                    }
                    if (osh_UK == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt1 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt1.WriteLine("Наименование УК не найдено в БД!" + " Ячейка: " + nom_cell_osh_UK);
                        nevtxt1.Close();
                    }
                    if (osh_UK_lic == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt2 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt2.WriteLine("Номер лицензии не соответствует наименованию УК!" + " Ячейка: " + nom_cell_osh_lic);
                        nevtxt2.Close();
                    }
                    if (osh_name_otch == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt2 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt2.WriteLine("Отсутствует наименование отчета!" + " Ячейка: " + cell_osh_name_otch);
                        nevtxt2.Close();
                    }
                    if (nomoshDate == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt2 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt2.WriteLine("Отсутствует дата!" + " Ячейка: " + nom_cell_date);
                        nevtxt2.Close();
                    }
                    if (osh_date_otch == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt2 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt2.WriteLine("Отсутствует дата отчета!" + " Ячейка: " + nom_cell_date_otch);
                        nevtxt2.Close();
                    }
                    if (osh_city == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt2 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt2.WriteLine("Отсутствует наименование города УК!" + " Ячейка: " + nom_cell_city);
                        nevtxt2.Close();
                    }
                    if (osh_lic_fond == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt2 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt2.WriteLine("Отсутствует номер лицензии фонда!" + " Ячейка: " + nom_cell_lic_fond);
                        nevtxt2.Close();
                    }
                    if (otmet_osh_uved == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt2 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt2.WriteLine("Отсутствует номер уведомления ЦБ!" + " Ячейка: " + nom_cell_uved);
                        nevtxt2.Close();
                    }
                    if (kol_osh > 0)
                    {
                        objWorkBook.Close();
                        objWorkExcel.Quit();
                        File.Move(starputy, novputy);                                                  // в ошибки
                    }

                    #endregion
                    objWorkExcel.Quit();
                }
                #endregion

                #region проверка на ошибки от ЦБ
                int nomer_name_otch_with_osh = 1;
                if (nazvaniefilebez.Contains("ошибки"))
                {
                    kodproverkinazvfile = 4;
                    Excel.Application objWorkExcel = new Excel.Application();                   //подключим excel
                    Excel.Workbook objWorkBook = objWorkExcel.Workbooks.Open(putfile, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    Excel.Worksheet objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1];      //получим 1 лист

                    #region //УК+лицензия
                    Excel.Range CellUK = objWorkSheet.get_Range("F17");
                    string nameUK = CellUK.Text.ToString();
                    Excel.Range CellLi = objWorkSheet.get_Range("F19");
                    string nameLi = CellLi.Text.ToString();

                    string nameUKBD, nameLiBD, nom_cell_osh_UK = "", nom_cell_osh_lic = "";
                    int nameIDpomoch = 1, nameID;
                    int nameIDUK = 0;
                    int n = 1, str_1 = 0;
                    CommandBD.CommandText = "SELECT * FROM Список_Управляющих_компаний";                     //выбор всего из таблицы UK
                    OleDbDataReader dr1 = CommandBD.ExecuteReader();                 //все начало делаться через путь выше
                    while (dr1.Read())
                    {
                        nameUKBD = dr1.GetString(3);                                 //считать 4 столбец  - название УК (начиная с 0)
                        nameLiBD = dr1.GetString(5);                                 //считать 6 столбец - лицензия
                        if (nameLiBD == nameLi)                                     //сравнение лицензии
                        {
                            nomoshUK_2 = 0;
                        }
                        else
                        {
                            nomoshUK_2 = 2;
                            nom_cell_osh_lic = "F19";
                        }
                        if (nameUKBD == nameUK)                                    //сравнение названия УК в БД и excel файле
                        {
                            nomoshUK_1 = 0;
                            nameID = dr1.GetInt32(2);                                //считать ID_UK
                            nameIDpomoch = nameID;
                            nameIDUK = nameID;
                            str_1 = 0;
                            break;
                        }
                        else
                        {
                            str_1++;
                        }
                        if (str_1 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_1 > 0)
                    {
                        nomoshUK_1 = 1;
                        nomoshUK_2 = 0;
                        nom_cell_osh_UK = "F17";
                    }
                    dr1.Close();
                    #endregion
                    #region //Дата
                    string nom_cell_date = "A8";
                    Excel.Range CellDate_otm = objWorkSheet.get_Range("A8");
                    string nameDate_otm = CellDate_otm.Text.ToString();
                    DateTime data_otmet = new DateTime();
                    if (nameDate_otm == "")
                    {
                        nomoshDate = 1;
                    }
                    else
                    {
                        nomoshDate = 0;
                        data_otmet = DateTime.Parse(nameDate_otm);
                    }
                    #endregion
                    #region //Дата отчета
                    string nom_cell_date_otch = "A26";
                    Excel.Range CellDate_otch = objWorkSheet.get_Range("A26");
                    string nameDate_otch = CellDate_otch.Text.ToString();
                    DateTime data_otch = new DateTime();
                    if (nameDate_otch == "")
                    {
                        osh_date_otch = 1;
                    }
                    else
                    {
                        osh_date_otch = 0;
                        data_otch = DateTime.Parse(nameDate_otch);
                    }
                    #endregion
                    #region //Фонд
                    Excel.Range CellFond = objWorkSheet.get_Range("F21");
                    string nameFond = CellFond.Text.ToString();

                    string nameFBD, nom_cell_osh_fond = "";
                    int nameIDF = 0, str_2 = 0;
                    int nameIDUKFond;
                    CommandBD.CommandText = "SELECT * FROM Список_фондов";
                    OleDbDataReader dr2 = CommandBD.ExecuteReader();
                    while (dr2.Read())
                    {
                        nameFBD = dr2.GetString(2);                                 //считать 4 столбец  - название Фонда (начиная с 0)
                        nameIDUKFond = dr2.GetInt32(0);
                        if (nameIDUKFond == nameIDpomoch)
                        {
                            nomoshF_2 = 0;
                        }
                        else
                        {
                            nomoshF_2 = 2;
                            nom_cell_osh_fond = "F21";
                        }
                        if (nameFBD == nameFond)                                    //сравнение названия фонда в БД и excel файле
                        {
                            nomoshF_1 = 0;
                            nameID = dr2.GetInt32(1);                                //считать ID
                            nameIDF = nameID;
                            str_2 = 0;
                            break;
                        }
                        else
                        {
                            str_2++;
                        }
                        if (str_2 > max_kol_strok_for_proverka) { break; }
                    }
                    if (str_2 > 0)
                    {
                        nomoshF_1 = 1;
                        nom_cell_osh_fond = "F21";
                    }
                    dr2.Close();
                    #endregion
                    #region //название отчета с ошибкой
                    string cell_name_with_osh = "A29";
                    Excel.Range Cell_nameot_with_osh = objWorkSheet.get_Range("A29");
                    string nameot_with_osh = Cell_nameot_with_osh.Text.ToString();
                    nomer_name_otch_with_osh = 0;
                    if (nameot_with_osh == "")
                    {
                        nomer_name_otch_with_osh = 1;
                    }
                    #endregion
                    #region //номер уведомления
                    string nom_cell_uved = "H8";
                    Excel.Range Cell_uved = objWorkSheet.get_Range("H8");
                    string nomer_uved = Cell_uved.Text.ToString();
                    otmet_osh_uved = 0;
                    if (nomer_uved == "")
                    {
                        otmet_osh_uved = 1;
                    }
                    #endregion
                    #region //город + заполнение лицензии (не проверка по БД!)
                    string nom_cell_city = "F18", nom_cell_lic_fond = "F22";
                    Excel.Range CellCity = objWorkSheet.get_Range("F18");
                    string nameCity = CellCity.Text.ToString();
                    Excel.Range Cell_lic_fond = objWorkSheet.get_Range("F22");
                    string name_lic_fond = Cell_lic_fond.Text.ToString();
                    osh_city = 0;
                    osh_lic_fond = 0;
                    if (nameCity == "")
                    {
                        osh_city = 1;
                    }
                    if (name_lic_fond == "")
                    {
                        osh_lic_fond = 1;
                    }
                    #endregion

                    #region //запись в БД
                    if (nomoshUK_1 == 0)
                    {
                        if (nomoshUK_2 == 0)
                        {
                            if (nomoshF_1 == 0)
                            {
                                if (nomoshF_2 == 0)
                                {
                                    if ((nomer_name_otch_with_osh == 0) && (nomoshDate == 0) && (otmet_osh_uved == 0) && (osh_date_otch == 0) && (osh_city == 0) && (osh_lic_fond == 0))
                                    {
                                        CommandBD.CommandText = "INSERT INTO [Ошибки_ЦБ] ([Дата], [Номер_уведомления], [Наименование_УК], [Наименование_фонда], [Дата_отчета], [Название_отчета]) VALUES ('" + data_otmet + "', '" + nomer_uved + "', '" + nameUK + "', '" + nameFond + "', '" + data_otch + "', '" + nameot_with_osh + "')";
                                        CommandBD.ExecuteNonQuery();
                                        conn.Close();
                                        objWorkBook.Close();
                                        objWorkExcel.Quit();
                                        File.Copy(starputy, putyarhiv_from_CB); //в архив ошибок от ЦБ
                                        File.Move(starputy, novputyarhiv);  //в архив
                                    }
                                }
                            }
                        }
                    }
                    #endregion

                    #region проверка на ошибки и указание их типа
                    int kol_osh = 0;

                    if (nomoshF_1 == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt3 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt3.WriteLine("Наименование Фонда не найдено в БД!" + " Ячейка: " + nom_cell_osh_fond);
                        nevtxt3.Close();
                    }
                    if (nomoshF_2 == 2)
                    {
                        kol_osh++;
                        StreamWriter nevtxt4 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt4.WriteLine("Наименование Фонда не соответствует наименованию УК!" + " Ячейка: " + nom_cell_osh_fond);
                        nevtxt4.Close();
                    }
                    if (nomoshUK_1 == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt1 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt1.WriteLine("Наименование УК не найдено в БД!" + " Ячейка: " + nom_cell_osh_UK);
                        nevtxt1.Close();
                    }
                    if (nomoshUK_2 == 2)
                    {
                        kol_osh++;
                        StreamWriter nevtxt2 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt2.WriteLine("Номер лицензии не соответствует наименованию УК!" + " Ячейка: " + nom_cell_osh_lic);
                        nevtxt2.Close();
                    }
                    if (nomer_name_otch_with_osh == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt2 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt2.WriteLine("Ошибка в наименовании отчета с ошибкой!" + " Ячейка: " + cell_name_with_osh);
                        nevtxt2.Close();
                    }
                    if (otmet_osh_uved == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt2 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt2.WriteLine("Отсутствует номер уведомления ЦБ!" + " Ячейка: " + nom_cell_uved);
                        nevtxt2.Close();
                    }
                    if (nomoshDate == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt2 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt2.WriteLine("Отсутствует дата!" + " Ячейка: " + nom_cell_date);
                        nevtxt2.Close();
                    }
                    if (osh_date_otch == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt2 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt2.WriteLine("Отсутствует дата отчета!" + " Ячейка: " + nom_cell_date_otch);
                        nevtxt2.Close();
                    }
                    if (osh_city == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt2 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt2.WriteLine("Отсутствует наименование города УК!" + " Ячейка: " + nom_cell_city);
                        nevtxt2.Close();
                    }
                    if (osh_lic_fond == 1)
                    {
                        kol_osh++;
                        StreamWriter nevtxt2 = File.AppendText(novputtxt);                          // и создаем ему файл ошибок с текстом
                        nevtxt2.WriteLine("Отсутствует номер лицензии фонда!" + " Ячейка: " + nom_cell_lic_fond);
                        nevtxt2.Close();
                    }
                    if (kol_osh > 0)
                    {
                        objWorkBook.Close();
                        objWorkExcel.Quit();
                        File.Move(starputy, novputy);                                                  // в ошибки
                    }

                    #endregion
                    objWorkExcel.Quit();
                }
                #endregion

                #region проверка на название файла
                if (kodproverkinazvfile != 1)
                {
                    if (kodproverkinazvfile != 2)
                    {
                        if (kodproverkinazvfile != 3)
                        {
                            if (kodproverkinazvfile != 4)
                            {
                                conn.Close();
                                File.Move(starputy, novputy);
                                string text_osh_name = "Ошибка в наименовании файла!";
                                Zapis_text(novputtxt, text_osh_name);
                            }
                        }
                    }
                }
                #endregion
            }
            //MessageBox.Show("Записано в БД!", "Уведомление");
        }
        #endregion
    }
}
