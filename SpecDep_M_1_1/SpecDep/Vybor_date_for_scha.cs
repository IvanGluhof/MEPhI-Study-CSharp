using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace SpecDep
{
    public partial class Vybor_date_for_scha : Form
    {
        public static string vybran_date = "";
        public static bool cancel_kod_date = false;

        public Vybor_date_for_scha()
        {
            InitializeComponent();
        }

        private void Vybor_date_for_scha_Load(object sender, EventArgs e)
        {
            monthCalendar1.Show();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //monthCalendar1.Show();
        }

        private void button_cancel_Click(object sender, EventArgs e)
        {
            DialogResult dialog = MessageBox.Show("Отмена формирования отчета!" + Environment.NewLine + "Продолжить?", "Уведомление", MessageBoxButtons.YesNo);
            if (dialog == DialogResult.Yes)
            {
                cancel_kod_date = true;
                this.Close();
            }
            if (dialog == DialogResult.No)
            {
                return;
            }
        }

        private void button_save_Click(object sender, EventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(textBox_date.Text))
            {
                vybran_date = textBox_date.Text;
                this.Close();
            }
            if (string.IsNullOrWhiteSpace(textBox_date.Text))
            {
                MessageBox.Show("Выберите предыдущую дату для формирования отчета!", "Нельзя продолжить без даты!");
            }
        }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {

        }

        private void textBox_date_DoubleClick(object sender, EventArgs e)
        {
        }

        private void monthCalendar1_MouseLeave(object sender, EventArgs e)
        {
        }

        private void monthCalendar1_DateSelected(object sender, DateRangeEventArgs e)
        {
            vybran_date = textBox_date.Text = monthCalendar1.SelectionStart.ToShortDateString().ToString();
        }
    }
}
