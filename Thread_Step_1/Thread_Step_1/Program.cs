using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Thread_Step_1
{
    class Program
    {

        static void WriteString(object _Data)
        {
            // для получения строки используем преобразование типов:
            // приводим переменную _Data к типу string и записываем
            // в переменную str_for_out
            string str_for_out = (string)_Data;
            // теперь поток 1 тысячу раз выведет полученную строку (свой номер)
            for (int i = 0; i <= 1000; i++)
            {
                Console.Write(str_for_out);
            }
        }

        static void Main(string[] args)
        {
            // создаём 4 потока, в качестве параметров передаём имя Выполняемой функции
            Thread th_1 = new Thread(WriteString);
            Thread th_2 = new Thread(WriteString);
            Thread th_3 = new Thread(WriteString);
            Thread th_4 = new Thread(WriteString);

            // расставляем приоритеты для потоков
            th_1.Priority = ThreadPriority.Highest;
            th_2.Priority = ThreadPriority.BelowNormal;
            th_3.Priority = ThreadPriority.Normal;
            th_4.Priority = ThreadPriority.Lowest;

            // запускаем каждый поток, в качестве параметра передаём номер потока
            th_1.Start("1");
            th_2.Start("2");
            th_3.Start("3");
            th_4.Start("4");

            Console.WriteLine("Все потоки запущен");

            // Ждём завершения каждого потока
            th_1.Join();
            th_2.Join();
            th_3.Join();
            th_4.Join();

            Console.ReadKey(); // прочитать символ (пока пользователь не нажмёт клавишу, программа не завершится (чтобы можно было успеть посмотреть результат)

        }
    }
}
