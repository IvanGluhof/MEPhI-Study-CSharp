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
using Calendar.NET;
using Excel = Microsoft.Office.Interop.Excel;

namespace SpecDep
{
    /// <summary>
    /// Класс для работы с базой данных.
    /// Здесь хранится всё, что связано ТОЛЬКО с работой в базе данных - выборкой, обновлением, изменением и прочим.
    /// </summary>
    class DataBase_FrameWork
    {
        static string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DataBaseMy.accdb";

        private HomeForm work_form;

        private string tablename;

        List<string> dates = new List<string>();
        List<string> values = new List<string>();

        List<string> tables_one = new List<string>();
        List<string> tables_two = new List<string>();

        public List<string> Dates
        {
            get { return dates; }
        }

        public List<string> Values
        {
            get { return values; }
        }

        public DataBase_FrameWork(HomeForm form)
        {
            work_form = form;

            // 1. Установка таблиц
            Set_Table_Lists();
            // 2. Подключение
            Connect_to_DB();
            // 3. Данные для событий в календаре
            //Read_Dates_For_Calendar();

            //4. Автозаполнение и выпадающие списки
            //FillDropDownList();
            Combo_collection_name_UK();
            Combo_collection_name_fond();
        }

        private void test()
        {

        }

        private void Combo_collection_name_UK()     //Названия УК
        {
            AutoCompleteStringCollection combo_collection_2 = new AutoCompleteStringCollection();
            OleDbCommand CommandBD = new OleDbCommand();                                      //команда, через которую все делается
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DataBaseMy.accdb";
            OleDbConnection conn = new OleDbConnection(connectionString);                       //новое подключение к БД
            CommandBD.Connection = conn;                                                      //соединение с бд
            conn.Open();
            CommandBD.CommandText = "SELECT Полное_наименование_УК FROM Список_Управляющих_компаний";
            OleDbDataReader dr2 = CommandBD.ExecuteReader();

            while (dr2.Read())
            {
                combo_collection_2.Add(dr2["Полное_наименование_УК"].ToString());
                work_form.comboBox_UK.Items.Add(dr2["Полное_наименование_УК"].ToString());
                work_form.comboBox_UK_otch.Items.Add(dr2["Полное_наименование_УК"].ToString());
            }
            dr2.Close();

            work_form.comboBox_UK.AutoCompleteCustomSource = combo_collection_2;
            work_form.comboBox_UK.AutoCompleteSource = AutoCompleteSource.CustomSource;
            work_form.comboBox_UK.AutoCompleteMode = AutoCompleteMode.SuggestAppend;

            work_form.comboBox_UK_otch.AutoCompleteCustomSource = combo_collection_2;
            work_form.comboBox_UK_otch.AutoCompleteSource = AutoCompleteSource.CustomSource;
            work_form.comboBox_UK_otch.AutoCompleteMode = AutoCompleteMode.SuggestAppend;

            conn.Close();
        }

        private void Combo_collection_name_fond()   //Названия фондов
        {
            AutoCompleteStringCollection combo_collection_3 = new AutoCompleteStringCollection();
            OleDbCommand CommandBD = new OleDbCommand();                                      //команда, через которую все делается
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DataBaseMy.accdb";
            OleDbConnection conn = new OleDbConnection(connectionString);                       //новое подключение к БД
            CommandBD.Connection = conn;                                                      //соединение с бд
            conn.Open();
            CommandBD.CommandText = "SELECT Полное_наименование_фонда FROM Список_фондов";
            OleDbDataReader dr3 = CommandBD.ExecuteReader();

            while (dr3.Read())
            {
                combo_collection_3.Add(dr3["Полное_наименование_фонда"].ToString());
                work_form.comboBox_Fond.Items.Add(dr3["Полное_наименование_фонда"].ToString());
                work_form.comboBox_Fond_otch.Items.Add(dr3["Полное_наименование_фонда"].ToString());
            }
            dr3.Close();

            work_form.comboBox_Fond.AutoCompleteCustomSource = combo_collection_3;
            work_form.comboBox_Fond.AutoCompleteSource = AutoCompleteSource.CustomSource;
            work_form.comboBox_Fond.AutoCompleteMode = AutoCompleteMode.SuggestAppend;

            work_form.comboBox_Fond_otch.AutoCompleteCustomSource = combo_collection_3;
            work_form.comboBox_Fond_otch.AutoCompleteSource = AutoCompleteSource.CustomSource;
            work_form.comboBox_Fond_otch.AutoCompleteMode = AutoCompleteMode.SuggestAppend;

            conn.Close();
        }

        /*protected internal void FillDropDownList()
        {
            try
            {
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    string strSql = "SELECT DISTINCT Управляющая_компания FROM [Согласия]";
                    OleDbDataAdapter adapter = new OleDbDataAdapter(new OleDbCommand(strSql, conn));
                    DataSet ds = new DataSet();
                    adapter.Fill(ds);
                    work_form.comboBox_UK.DataSource = ds.Tables[0];
                    work_form.comboBox_UK.DisplayMember = "Управляющая_компания";
                    work_form.comboBox_UK.ValueMember = "Управляющая_компания";

                    strSql = "SELECT DISTINCT Фонд FROM [Согласия]";
                    adapter = new OleDbDataAdapter(new OleDbCommand(strSql, conn));
                    ds = new DataSet();
                    adapter.Fill(ds);
                    work_form.comboBox_Fond.DataSource = ds.Tables[0];
                    work_form.comboBox_Fond.DisplayMember = "Фонд";
                    work_form.comboBox_Fond.ValueMember = "Фонд";
                }
            }
            catch (Exception ex)
            {

            }

        }*/

        private void Set_Table_Lists()
        {
            tables_one.Add("Согласия");
            tables_one.Add("Платежные_поручения");

            tables_two.Add("Выписка_день");
            tables_two.Add("Данные_день_выписки");
            tables_two.Add("Данные_день_СЧА");
            tables_two.Add("Контрагент");
            tables_two.Add("Ответственное_лицо");
            tables_two.Add("Пайщики_фонда");
            tables_two.Add("Предоставляет_паи");
            //tables_two.Add("Согласие_плат_пор");
            tables_two.Add("Список_банков");
            tables_two.Add("Список_Управляющих_компаний");
            tables_two.Add("Список_фондов");
            tables_two.Add("Справочник_валют");
            tables_two.Add("Справочник_оснований");
            tables_two.Add("Справочник_сделок");
            tables_two.Add("Справочник_стран");
            tables_two.Add("СЧА_день");
            tables_two.Add("Ошибки_ЦБ");

        }

        /// <summary>
        /// Подключение к БД
        /// </summary>
        /// <param name="comboBox_view"></param>
        private void Connect_to_DB()
        {
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                conn.Open();

                //ограничение - чтобы не выводил системные таблицы
                string[] restrictions = new string[4];
                DataTable userTables = null;

                foreach (string table in tables_one)
                {
                    restrictions[2] = table;
                    
                    userTables = conn.GetSchema("Tables", restrictions);
                    work_form.comboBox_view.Items.AddRange(new object[] { userTables.Rows[0][2].ToString() });
                }

                foreach (string table in tables_two)
                {
                    restrictions[2] = table;

                    userTables = conn.GetSchema("Tables", restrictions);
                    work_form.comboBox_sp.Items.AddRange(new object[] { userTables.Rows[0][2].ToString() });
                }

                work_form.comboBox_view.SelectedIndex = 0; // По умолчанию всегда выбрана первая таблица
                work_form.comboBox_sp.SelectedIndex = 0; // По умолчанию всегда выбрана первая таблица
                View_Table(); // И сразу загружает
                View_Table_sp();
            }
        }

        /// <summary>
        /// Даты для календаря
        /// </summary>
        protected internal void Read_Dates_For_Calendar()
        {
            dates.Clear();
            values.Clear();
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                OleDbCommand CommandBD = new OleDbCommand();
                conn.Open();

                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM СЧА_день", conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "СЧА_день");
                
                foreach (DataRow row in ds.Tables["СЧА_день"].Rows)
                {
                    dates.Add(row["Дата_отчета"].ToString());
                }

                foreach (DataRow row in ds.Tables["СЧА_день"].Rows)
                {
                    values.Add("СЧА: " + row["СЧА"].ToString() + Environment.NewLine + "Паев: " + 
                        row["Кол_паев"].ToString() + Environment.NewLine + "Цена пая: " + row["Цена_пая"].ToString());
                }
            }
        }

        protected internal void Refresh_Dates_For_Calendar(string uk, string fond)
        {
            dates.Clear();
            values.Clear();
            using (OleDbConnection conn = new OleDbConnection(connectionString))
            {
                OleDbCommand CommandBD = new OleDbCommand();
                conn.Open();

                OleDbDataAdapter da = new OleDbDataAdapter("SELECT * FROM СЧА_день WHERE УК LIKE '%" + uk + "' AND ФОНД LIKE '%" + fond + "'", conn);
                DataSet ds = new DataSet();
                da.Fill(ds, "СЧА_день");

                foreach (DataRow row in ds.Tables["СЧА_день"].Rows)
                {
                    dates.Add(row["Дата_отчета"].ToString());
                }

                foreach (DataRow row in ds.Tables["СЧА_день"].Rows)
                {
                    values.Add("СЧА: " + row["СЧА"].ToString() + Environment.NewLine + "Паев: " +
                        row["Кол_паев"].ToString() + Environment.NewLine + "Цена пая: " + row["Цена_пая"].ToString());
                }
            }
        }

        /// <summary>
        /// Просмотр таблицы
        /// </summary>
        protected internal void View_Table()
        {
            if (!string.IsNullOrWhiteSpace(Convert.ToString(work_form.comboBox_view.SelectedItem)))
            {
                tablename = Convert.ToString(work_form.comboBox_view.SelectedItem);
                string vybor_iz_bd = "SELECT * FROM " + tablename;
                
                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    OleDbDataAdapter dataadapter = new OleDbDataAdapter(vybor_iz_bd, conn);
                    DataSet ds = new DataSet();

                    dataadapter.Fill(ds, tablename);
                    conn.Close();
                    work_form.dataGridView_DB.DataSource = ds;
                    work_form.dataGridView_DB.DataMember = tablename;
                }
                //перенос текста в ячейке datagridview
                work_form.dataGridView_DB.DefaultCellStyle.WrapMode = DataGridViewTriState.True;

            }
            else
            {
                MessageBox.Show("Не выбрана таблица для просмотра!");
            }
        }
        protected internal void View_Table_sp()
        {
            if (!string.IsNullOrWhiteSpace(Convert.ToString(work_form.comboBox_sp.SelectedItem)))
            {
                tablename = Convert.ToString(work_form.comboBox_sp.SelectedItem);
                string vybor_iz_bd = "SELECT * FROM " + tablename;

                using (OleDbConnection conn = new OleDbConnection(connectionString))
                {
                    conn.Open();
                    OleDbDataAdapter dataadapter = new OleDbDataAdapter(vybor_iz_bd, conn);
                    DataSet ds = new DataSet();

                    dataadapter.Fill(ds, tablename);
                    conn.Close();
                    work_form.dataGridView_sp.DataSource = ds;
                    work_form.dataGridView_sp.DataMember = tablename;
                }
                //перенос текста в ячейке datagridview
                work_form.dataGridView_DB.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            }
            else
            {
                MessageBox.Show("Не выбрана таблица для просмотра!");
            }
        }

        /// <summary>
        /// Сохранение таблицы
        /// </summary>
        protected internal void Save_to_BD()
        {
            try
            {
                string queryString = "SELECT * FROM " + work_form.comboBox_view.Text;
                using (OleDbConnection connection =
                           new OleDbConnection(connectionString))
                {
                    connection.Open();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = new OleDbCommand(queryString, connection);
                    OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter);

                    // Code to modify data in the DataSet here.
                    DataSet ds = work_form.dataGridView_DB.DataSource as DataSet;
                    // Without the OleDbCommandBuilder, this line would fail.
                    adapter.UpdateCommand = builder.GetUpdateCommand();
                    adapter.Update(ds, work_form.comboBox_view.Text);
                }
                //FillDropDownList();
                MessageBox.Show("Сохранение успешно");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка: \r\n" + ex.Message);
            }
        }
        protected internal void Save_to_BD_sp()
        {
            try
            {
                string queryString = "SELECT * FROM " + work_form.comboBox_sp.Text;
                using (OleDbConnection connection =
                           new OleDbConnection(connectionString))
                {
                    connection.Open();
                    OleDbDataAdapter adapter = new OleDbDataAdapter();
                    adapter.SelectCommand = new OleDbCommand(queryString, connection);
                    OleDbCommandBuilder builder = new OleDbCommandBuilder(adapter);

                    // Code to modify data in the DataSet here.
                    DataSet ds = work_form.dataGridView_sp.DataSource as DataSet;
                    // Without the OleDbCommandBuilder, this line would fail.
                    adapter.UpdateCommand = builder.GetUpdateCommand();
                    adapter.Update(ds, work_form.comboBox_sp.Text);
                }
                MessageBox.Show("Сохранение успешно");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка - " + ex.Message);
            }
        }

        protected internal void Delete_Save_Row()
        {
            DialogResult result = MessageBox.Show("Вы действительно хотите удалить строки?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                foreach (DataGridViewRow item in work_form.dataGridView_DB.SelectedRows)
                {
                    work_form.dataGridView_DB.Rows.RemoveAt(item.Index);
                }
                Save_to_BD();
            }
        }
        protected internal void Delete_Save_Row_sp()
        {
            DialogResult result = MessageBox.Show("Вы действительно хотите удалить строки?", "Внимание", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                foreach (DataGridViewRow item in work_form.dataGridView_sp.SelectedRows)
                {
                    work_form.dataGridView_sp.Rows.RemoveAt(item.Index);
                }
                Save_to_BD();
            }
        }

        //Это метод. Подробнее:
        //1)private - уровень доступа (можешь вызывать только в своём классе = своей форме)
        //2)string - тип возвращаемой функции. То есть какой тип переменной метод вернет, то есть сохранит и передаст, после выполнения
        //3)В скобочках - тип и название переменных, которые мы передаем методу и дальше внутри него используем
        private string Query_Builder(string column_to_find, string value, string query_add, bool first)
        {
            int column_count = work_form.dataGridView_DB.Columns.Count; //Количество столбцов в таблице
            //Цикл для поиска
            //Тут важно уменьшить количество столбцов на 1. Иначе будет выход за пределы цикла
            for (int i = 0; i <= column_count - 1; i++)
            {
                if (work_form.dataGridView_DB.Columns[i].HeaderText == column_to_find) //Проверка условия - на каждом прогоне проверяем, равно ли имя столбца тому, что мы ищем
                {
                    if (first == true)
                    {
                        //MessageBox.Show(i.ToString() + " " + dataGridView1.Columns[i].HeaderText); //Debug - для тестирования
                        query_add = query_add + " WHERE " + column_to_find + " LIKE '%" + value + "%'"; //Если нашли, то к SQL запросу добавляем нужные параметры поиска
                    }
                    else
                    {
                        query_add = query_add + " AND " + column_to_find + " LIKE '%" + value + "%'";
                    }
                }
            }
            return query_add; //А это возврат. Означает завершение работы метода и сохранения информации о переменной query_add. Возврат должен быть всегда
        }

        /// <summary>
        /// Фильтр
        /// </summary>
        protected internal void Do_Filter()
        {
            bool first = false;
            bool builder = false;
            if (!string.IsNullOrWhiteSpace(Convert.ToString(work_form.comboBox_view.SelectedItem)))
            {
                string table_name = Convert.ToString(work_form.comboBox_view.SelectedItem); //Название таблицы - берется из ComboBox'a
                string query_to_refresh = "SELECT * FROM " + table_name; //Выборка всей таблицы
                //Названия столблцов в базе данных должно быть строго таким же, как в текст.боксах, что над полями для поиска. Можно изменить это, но это непросто
                //Серия проверок - если поисковые строки заполнены, то в SQL запрос добавляются параметры для выборки
                //Тут конечно можно поумнее сделать, но не буду грузить - так зато всё понятно
                if (!string.IsNullOrWhiteSpace(work_form.textBox_Data.Text)) //! в C# это отрицание => проверяем, НЕ пуста ли строка. Если НЕ пустая - то выполняется if
                {
                    //Метод для проверки и изменения SQL запроса
                    //Ему передаём данные - что введено для поиска, и, грубо говоря, название столбца для поиска
                    //Результатом будем либо дополненный SQL запрос, либо без изменений
                    if (first == false) first = true;
                    query_to_refresh = Query_Builder(work_form.textBoxlabel_Data.Text, work_form.textBox_Data.Text, query_to_refresh, first);
                }
                if (!string.IsNullOrWhiteSpace(work_form.comboBox_UK.Text))
                {
                    if (first == false) { first = true; query_to_refresh = Query_Builder(work_form.textBoxlabel_UK.Text, work_form.comboBox_UK.Text, query_to_refresh, first); }
                    else if (first == true) { query_to_refresh = Query_Builder(work_form.textBoxlabel_UK.Text, work_form.comboBox_UK.Text, query_to_refresh, builder); }
                }
                if (!string.IsNullOrWhiteSpace(work_form.comboBox_Fond.Text))
                {
                    if (first == false) { first = true; query_to_refresh = Query_Builder(work_form.textBoxlabel_Fond.Text, work_form.comboBox_Fond.Text, query_to_refresh, first); }
                    else if (first == true) { query_to_refresh = Query_Builder(work_form.textBoxlabel_Fond.Text, work_form.comboBox_Fond.Text, query_to_refresh, builder); }
                }
                if (!string.IsNullOrWhiteSpace(work_form.textBox_Naznach.Text))
                {
                    if (first == false) { first = true; query_to_refresh = Query_Builder(work_form.textBoxlabel_Naznach.Text, work_form.textBox_Naznach.Text, query_to_refresh, first); }
                    else if (first == true) { query_to_refresh = Query_Builder(work_form.textBoxlabel_Naznach.Text, work_form.textBox_Naznach.Text, query_to_refresh, builder); }
                }

                if ((string.IsNullOrWhiteSpace(work_form.textBox_Data.Text)) && (string.IsNullOrWhiteSpace(work_form.comboBox_UK.Text)) && (string.IsNullOrWhiteSpace(work_form.comboBox_Fond.Text)) && (string.IsNullOrWhiteSpace(work_form.textBox_Naznach.Text)))
                {
                    MessageBox.Show("Ни одно из полей для поиска не заполнено");
                }
                else
                {
                    using (OleDbConnection connection = new OleDbConnection(connectionString))
                    {
                        OleDbDataAdapter dataadapter = new OleDbDataAdapter(query_to_refresh, connection);
                        DataSet ds = new DataSet();
                        connection.Open();
                        dataadapter.Fill(ds, table_name);
                        connection.Close();
                        work_form.dataGridView_DB.DataSource = ds;
                    }                  
                }
            }
            else
            {
                MessageBox.Show("Не выбрана таблица для поиска!");
            }
        }

        /// <summary>
        /// Согласия
        /// </summary>
        /// <returns></returns>
        protected internal string Load_Sogl()
        {         
            using (OleDbConnection conn = new OleDbConnection(connectionString)) //новое подключение к БД)
            {
                OleDbCommand CommandBD = new OleDbCommand(); //команда, через которую все делается
                CommandBD.Connection = conn;//соединение с бд
                conn.Open();
                CommandBD.CommandText = "SELECT * FROM Согласия WHERE [Отметка_согласования] = @Q1";
                CommandBD.Parameters.Add("@Q1", OleDbType.Boolean).Value = false;
                OleDbDataReader dr1 = CommandBD.ExecuteReader();

                int kolich = 0;
                while (dr1.Read())
                {
                    if (Convert.ToBoolean(CommandBD.Parameters["@Q1"].Value) == false)
                    {
                        kolich++;
                    }
                }
                string result = Convert.ToString(kolich);
                return result;
            }
        }

        /// <summary>
        /// Платежные поручения
        /// </summary>
        /// <returns></returns>
        protected internal string Load_plat_por()
        {
            using (OleDbConnection conn = new OleDbConnection(connectionString)) //новое подключение к БД)
            {
                OleDbCommand CommandBD = new OleDbCommand();                                      //команда, через которую все делается
                CommandBD.Connection = conn;                                                      //соединение с бд
                conn.Open();
                CommandBD.CommandText = "SELECT * FROM Платежные_поручения WHERE [Отметка_проверки] = @V1";
                CommandBD.Parameters.Add("@V1", OleDbType.Boolean).Value = false;

                OleDbDataReader dr1 = CommandBD.ExecuteReader();

                int kolich = 0;
                while (dr1.Read())
                {
                    if (Convert.ToBoolean(CommandBD.Parameters["@V1"].Value) == false)
                    {
                        kolich++;
                    }
                }
                string result = Convert.ToString(kolich);
                return result;
            }
        }
        
        /// <summary>
        /// Ошибки от ЦБ
        /// </summary>
        /// <returns></returns>
        protected internal string Load_osh_CB()
        {
            OleDbCommand CommandBD_3 = new OleDbCommand();                                      //команда, через которую все делается
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DataBaseMy.accdb";
            OleDbConnection conn = new OleDbConnection(connectionString);                       //новое подключение к БД
            CommandBD_3.Connection = conn;                                                      //соединение с бд
            conn.Open();
            CommandBD_3.CommandText = "SELECT * FROM Ошибки_ЦБ WHERE [Отметка_исправлено] = @W1";
            CommandBD_3.Parameters.Add("@W1", OleDbType.Boolean).Value = false;
            OleDbDataReader dr1 = CommandBD_3.ExecuteReader();

            int kolich = 0;
            while (dr1.Read())
            {
                if (Convert.ToBoolean(CommandBD_3.Parameters["@W1"].Value) == false)
                {
                    kolich++;
                }
            }
            string result = Convert.ToString(kolich);
            conn.Close();
            return result;
        }
    }
}
