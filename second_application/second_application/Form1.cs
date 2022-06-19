using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace second_application
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        Image MemForImage;
        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
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

        // обработка кнопки меню "загрузка - в формате jpg" 
        private void вФорматеJPGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadImage(true);
        }

        // обработка кнопки меню "загрузка - в формате png" 
        private void вФорматеPNGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            LoadImage(false);
        }

        // обработка кнопки №2 на панели 
        private void toolStripButton2_Click(object sender, EventArgs e)
        {
            LoadImage(true);
        }

        // обработка кнопки №3 на панели 
        private void toolStripButton3_Click(object sender, EventArgs e)
        {
            LoadImage(false);
        }

        // функция загрузки изображения 
        private void LoadImage(bool jpg)
        {
            // директория, которая будет выбрана как начальная в окне для выбора файла
            openFileDialog1.InitialDirectory = "c:";

            // если будем выбирать jpg файлы
            if(jpg)
            {
                // устанавливаем формат файлов для загрузки - jpg
                openFileDialog1.Filter = "image (JPEG) files (*.jpg)|*.jpg|All files (*.*)|*.*";
            }
            else
            {
                // устанавливаем формат файлов для загрузки - png
                openFileDialog1.Filter = "image (PNG) files (*.png)|*.png|All files (*.*)|*.*";
            }

            // если открытие окна выбора файла завершилось выбором файла и нажатием кнопки ОК
            if(openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try // безопасная попытка
                {
                    // пытаемся загрузить файл с именем openFileDialog1.FileName - выбранный пользователем файл.
                    MemForImage = Image.FromFile(openFileDialog1.FileName);
                    // устанавливаем картинку в поле элемента PictureBox
                    pictureBox1.Image = MemForImage;
                }
                catch (Exception ex) // если попытка загрузки не удалась
                {
                    // выводим сообщение с причиной ошибки
                    MessageBox.Show("Не удалось загрузить файл: " + ex.Message);
                }
            }
           
        }

        // кнопка активации дополнительного диалогового окна 
        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            // создаём новый экземпляр класса Preview
            // отвечающий за работу с нашей доп - ой формой
            // в качестве параметра мы передаем наше загруженное изображение
            Form PreView = new Preview(MemForImage);
            // затем мы вызовем диалоговое окно
            PreView.ShowDialog();
        }
    }
}
