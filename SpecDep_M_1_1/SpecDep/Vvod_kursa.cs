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
    public partial class Vvod_kursa : Form
    {
        public static string vybran_val = "";
        public static decimal vvod_kurs = 0;
        public static bool cancel_kod = false;
        public Vvod_kursa()
        {
            InitializeComponent();
            //выбор валют
            Combo_collection_val();
        }

        private void Vvod_kursa_Load(object sender, EventArgs e)
        {

        }

        private void comboBox_vybor_val_SelectedIndexChanged(object sender, EventArgs e)
        {

        }

        private void textBox_vvod_kursa_TextChanged(object sender, EventArgs e)
        {
            
        }

        private void button_cancel_Click(object sender, EventArgs e)
        {
            DialogResult dialog = MessageBox.Show("Отмена формирования отчета!" + Environment.NewLine + "Продолжить?", "Уведомление", MessageBoxButtons.YesNo);
            if (dialog == DialogResult.Yes)
            {
                cancel_kod = true;
                this.Close();
            }
            if (dialog == DialogResult.No)
            {
                return;
            }
        }

        private void button_save_Click(object sender, EventArgs e)
        {
            if ((!string.IsNullOrWhiteSpace(comboBox_vybor_val.Text)) && (!string.IsNullOrWhiteSpace(textBox_vvod_kursa.Text)))
            {
                vybran_val = comboBox_vybor_val.Text.Remove(3);
                vvod_kurs = Convert.ToDecimal(textBox_vvod_kursa.Text);
                this.Close();
            }
            if ((string.IsNullOrWhiteSpace(comboBox_vybor_val.Text)) || (string.IsNullOrWhiteSpace(textBox_vvod_kursa.Text)))
            {
                MessageBox.Show("Выберите  вид  валюты  и  введите  курс!", "Нельзя продолжить без курса валюты!");
            }
        }
        private void Combo_collection_val()     //Названия валют
        {
            AutoCompleteStringCollection combo_collection_2 = new AutoCompleteStringCollection();
            OleDbCommand CommandBD = new OleDbCommand();                                      //команда, через которую все делается
            string connectionString = @"Provider=Microsoft.ACE.OLEDB.12.0;Data Source=DataBaseMy.accdb";
            OleDbConnection conn = new OleDbConnection(connectionString);                       //новое подключение к БД
            CommandBD.Connection = conn;                                                      //соединение с бд
            conn.Open();
            CommandBD.CommandText = "SELECT Название, Описание FROM Справочник_валют";
            OleDbDataReader dr2 = CommandBD.ExecuteReader();
            while (dr2.Read())
            {
                combo_collection_2.Add(dr2["Название"].ToString() + ":" + dr2["Описание"].ToString());
                comboBox_vybor_val.Items.Add(dr2["Название"].ToString() + ":" + dr2["Описание"].ToString());
            }
            dr2.Close();
            //comboBox_vybor_val.AutoCompleteCustomSource = combo_collection_2;
            //comboBox_vybor_val.AutoCompleteSource = AutoCompleteSource.CustomSource;
            //comboBox_vybor_val.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            conn.Close();
        }
    }
}
