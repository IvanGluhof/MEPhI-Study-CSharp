using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Sorting
{
    class Random_Array
    {
        int[] value;

        public Random_Array(int size)
        {
            this.value = new int[size];
            Random rnd = new Random();

            for(int i = 0; i < size; i++)
            {
                value[i] = rnd.Next(100);
            }
            Console.WriteLine("Массив случайно сгенерированных чисел. Размер: " + this.Size());
            this.PrintArray();
        }

        public void PrintArray()
        {
            foreach(int v in value)
            {
                Console.Write(v + " ");
            }
            Console.WriteLine();
            Console.ReadLine();
        }

        public int Size()
        {
            int result = value.Length;
            return result;
        }

        // Сортировка простыми вставками
        public int[] Insert_Sort()
        {
            /*
            Все элементы условно разделяются на готовую
            последовательность a1 ... ai-1 и входную ai ... an. Hа
            каждом шаге, начиная с i=2 и увеличивая i на 1, берем i-й
            элемент входной последовательности и вставляем его на
            нужное место в готовую.
            */
            int temp, i, j;

            /***********************
             * сортируем a[0..n] *
             ***********************/
            for (i = 1; i < this.Size(); i++)
            {
                temp = value[i];

                /* Сдвигаем элементы вниз, пока */
                /*  не найдем место вставки.    */
                for (j = i - 1; j >= 0 && value[j] > temp; j--)
                    value[j + 1] = value[j];

                /* вставка */
                value[j + 1] = temp;
            }

            Console.WriteLine("Массив случайно сгенерированных чисел после сортировки вставкой");
            this.PrintArray();
            return value;
        }

        public void Quick_Sort(int l, int r)
        {
            int temp;
            int x = this.value[l + (r - l) / 2];
            //запись эквивалентна (l+r)/2,
            //но не вызввает переполнения на больших данных
            int i = l;
            int j = r;
            //код в while обычно выносят в процедуру particle
            while (i <= j)
            {
                while (this.value[i] < x) i++;
                while (this.value[j] > x) j--;
                if (i <= j)
                {
                    temp = this.value[i];
                    this.value[i] = this.value[j];
                    this.value[j] = temp;
                    i++;
                    j--;
                }
            }
            if (i < r)
                Quick_Sort(i, r);

            if (l < j)
                Quick_Sort(l, j);
        }
    }

    class Program
    {
        static void Main(string[] args)
        {
            Random_Array arr = new Random_Array(10);
            //arr.Insert_Sort();

            arr.Quick_Sort(0, arr.Size() - 1);

            arr.PrintArray();
        }
    }
}
