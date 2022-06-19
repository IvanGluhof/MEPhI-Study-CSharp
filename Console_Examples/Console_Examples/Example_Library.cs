using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace Console_Examples
{
    partial class Program
    {
        private static void Evoke_Example_Collection()
        {
            bool exit = false;

            Console.WriteLine("EXAMPLE LIBRARY");
            Console.WriteLine("Press number to call an exmaple");
            Console.WriteLine("To finish type Exit or exit");
            
            while (exit == false)
            {
                var key = Console.ReadLine();
                switch (key)
                {
                    case "1":
                        {
                            Call_Example_1();
                            break;
                        }
                    case "2":
                        {
                            Call_Example_2();
                            break;
                        }
                    case "3":
                        {
                            Call_Example_3();
                            break;
                        }
                    case "4":
                        {
                            Call_Example_4();
                            break;
                        }
                    case "5":
                        {
                            Call_Example_5();
                            break;
                        }
                    case "6":
                        {
                            Call_Example_6();
                            break;
                        }
                    case "7":
                        {
                            Call_Example_7();
                            break;
                        }
                    case "8":
                        {
                            Call_Example_8();
                            break;
                        }
                    case "9":
                        {
                            Call_Example_9();
                            break;
                        }
                    case "10":
                        {
                            Call_Example_10();
                            break;
                        }
                    case "Exit":
                    case "exit":
                        {
                            exit = true;
                            break;
                        }
                }
            }
        }

        #region EXAMPLE 1
        private static void Call_Example_1()
        {
            // Форматируем шапку программы
            Console.BackgroundColor = ConsoleColor.Green;
            Console.ForegroundColor = ConsoleColor.Black;
            Console.WriteLine("*****EXAMPLE 1******");
            Console.WriteLine("********************");
            Console.WriteLine("**** Мой проект ****");
            Console.WriteLine("********************");
            // Основная программа
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine();
            Console.WriteLine("Hello, World!");
            Console.BackgroundColor = ConsoleColor.Black;
            Console.ForegroundColor = ConsoleColor.Gray;
        }
        #endregion

        #region EXAMPLE 2
        private static void Call_Example_2()
        {
            var name = "Alex Erohin";
            var age = 26;
            var isProgrammer = true;

            // Определяем тип переменных
            Type nameType = name.GetType();
            Type ageType = age.GetType();
            Type isProgrammerType = isProgrammer.GetType();

            // Выводим в консоль результаты
            Console.WriteLine("Тип name: {0}", nameType);
            Console.WriteLine("Тип age {0}", ageType);
            Console.WriteLine("Тип isProgrammer {0}", isProgrammerType);
        }
        #endregion

        #region EXAMPLE 3
        private static void Call_Example_3()
        {
            for (int i = 0; i < 10; i++)
            {
                Console.Write(" {0}", i);
            } // здесь i покидает область видимости
            Console.WriteLine();

            // мы можем вновь объявить i
            for (int i = 0; i > -10; i--)
            {
                Console.Write(" {0}", i);
            } // i снова покидает область видимости
            Console.WriteLine();
            //var j = i * i;  // данный код не выполнится, т.к  i не определена в текущем контексте
        }
        #endregion

        #region Example 4
        private static void Call_Example_4()
        {
            long result;
            const long km = 149800000; // расстояние в км.

            result = km * 1000 * 100;
            Console.WriteLine(result);
        }
        #endregion

        #region Example 5
        private static void Call_Example_5()
        {
            // *** Расчет стоимости капиталовложения с ***
            // *** фиксированной нормой прибыли***
            decimal money, percent;
            int i;
            const byte years = 15;

            money = 1000.0m;
            percent = 0.045m;

            for (i = 1; i <= years; i++)
            {
                money *= 1 + percent;
            }

            Console.WriteLine("Общий доход за {0} лет: {1} $$", years, money);
        }
        #endregion

        #region Example 6
        private static void Call_Example_6()
        {
            // Используем перенос строки
            Console.WriteLine("Первая строка\nВторая строка\nТретья строка\n");

            // Используем вертикальную табуляцию
            Console.WriteLine("Первый столбец \v Второй столбец \v Третий столбец \n");

            // Используем горизонтальную табуляцию
            Console.WriteLine("One\tTwo\tThree");
            Console.WriteLine("Four\tFive\tSix\n");

            //Вставляем кавычки
            Console.WriteLine("\"Зачем?\", - спросил он");
        }
        #endregion
        #region Example 7

        private static void Call_Example_7()
        {
            int i1 = 455, i2 = 84500;
            decimal dec = 7.98845m;

            // Приводим два числа типа int
            // к типу short
            Console.WriteLine((short)i1);
            Console.WriteLine((short)i2);

            // Приводим число типа decimal
            // к типу int
            Console.WriteLine((int)dec);
        }
        #endregion

        #region Example 8
        private static void Call_Example_8()
        {
            byte var1 = 250;
            byte var2 = 150;
            try
            {
                byte sum = checked((byte)(var1 + var2));
                Console.WriteLine("Сумма: {0}", sum);
            }
            catch (OverflowException ex)
            {
                Console.WriteLine(ex.Message);
                Console.ReadLine();
            }
        }
        #endregion

        #region Example 9
        private static void Call_Example_9()
        {
            int num1, num2;
            float f1, f2;

            num1 = 13 / 3;
            num2 = 13 % 3;

            f1 = 13.0f / 3.0f;
            f2 = 13.0f % 3.0f;

            Console.WriteLine("Результат и остаток от деления 13 на 3: {0} __ {1}", num1, num2);
            Console.WriteLine("Результат деления 13.0 на 3.0: {0:#.###} {1}", f1, f2);
        }
        #endregion

        #region Example 10
        private static void Call_Example_10()
        {
            short d = 1;

            for (byte i = 0; i < 10; i++)
                Console.Write(i + d++ + "\t");

            Console.WriteLine();
            d = 1;

            for (byte i = 0; i < 10; i++)
                Console.Write(i + ++d + "\t");
        }
        #endregion

        #region Example
        private static void Call_Example_()
        {

        }
        #endregion
    }
}