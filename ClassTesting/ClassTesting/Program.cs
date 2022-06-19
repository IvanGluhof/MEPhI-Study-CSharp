using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ClassTesting
{
    class Program
    {
        static void Main(string[] args)
        {
            // переменная, которая будет хранить команду пользователя
            string user_command = "";

            // бесконечный цикл
            bool Infinity = true;

            // пустой (равный null) экземпляр класса Man
            Man SomeMan = null;

            while (Infinity == true) // пока Infinity = true
            {
                // приглашение ввести команду
                System.Console.WriteLine("Пожалуйста, введите команду");

                // получение строки (команды) от пользователя
                user_command = System.Console.ReadLine();

                // обработка команды с помощью оператора ветвления
                switch (user_command)
                {
                    case "create_man":
                        {
                            // просим ввести имя человека
                            System.Console.WriteLine("Пожалуйста, введите имя создаваемого человека");

                            // получаем строку введенную пользователем
                            user_command = System.Console.ReadLine();

                            // создаем новый объект в памяти, в качестве параметра
                            // передаем имя человека
                            SomeMan = new Man(user_command);

                            // сообщаем о создании
                            System.Console.WriteLine("Человек успешно создан");

                            break;
                        }

                    case "kill_man":
                        {
                            // проверяем, чтоб объект существует в памяти
                            if (SomeMan !=null)
                            {
                                // вызываем функцию смерти
                                SomeMan.Kill();
                            }
                            
                            break;
                        }

                    case "talk":
                        {
                            // проверяем, что объект существует в памяти
                            if (SomeMan != null)
                            {
                                //вызываем функцию разговора
                                SomeMan.Talk();
                            }
                            else
                            {
                                System.Console.WriteLine("Человек не создан. Команда не может быть выполнена");
                            }
                            break;
                        }

                    case "go":
                        {
                            // проверяем, что объект существует в памяти
                            if (SomeMan != null)
                            {
                                //вызываем функцию разговора
                                SomeMan.Go();
                            }
                            else
                            {
                                System.Console.WriteLine("Человек не создан. Команда не может быть выполнена");
                            }
                            break;
                        }

                    // если user_command содержит строку exit
                    case "exit":
                        {
                            Infinity = false;
                            // теперь цикл завершиться, и программа завершит свое выполнение
                            break;
                        }

                    // если user_command содержит строку help
                    case "help":
                        {
                            System.Console.WriteLine("Список команд");
                            System.Console.WriteLine("---");

                            System.Console.WriteLine("create_man : команда создает человека, (экземпляр класса Man)");
                            System.Console.WriteLine("kill_man : команда убивает человека");
                            System.Console.WriteLine("talk : команда заставляет человека говорить (если создан экземпляр класса)");
                            System.Console.WriteLine("go : команда заставляет человека идти (если создан экземпляр класса)");

                            System.Console.WriteLine("---");
                            System.Console.WriteLine("---");
                            break;
                        }

                    // если команду определить не удалось
                    default:
                        {
                            System.Console.WriteLine("Ваша команда не определена, пожалуйста, повторите снова");
                            System.Console.WriteLine("Для вывода списка команд введите команду help");
                            System.Console.WriteLine("Для завершения программы введите команду exit");
                            break;
                        }
                }
            }
        }
    }

    public class Man
    {
        // конструктор класса (данная функция вызывается
        // при создании нового экземпляра класса)

        public Man(string _name)
        {
            // получаем имя человека из входного параметра
            // конструктора класса
            Name = _name;
            // экземпляр жив
            isLife = true;

            

            // генерируем возраст от 15 до 50
            Age = (uint)rnd.Next(15, 51);
            // и здоровье, от 10 до 100%
            Health = (uint)rnd.Next(10, 101);
        }

        // экземпляр класса Random
        // для генерации случайных чисел
        private Random rnd = new Random();

        // закрытые члены, которые нельзя изменить
        // извне класса

        // строка, содержащая имя
        private string Name;
        
        // беззнаковое целое число = возраст
        private uint Age;

        // беззнаковое целое число = уровень здоровья
        private uint Health;

        // булево, означающее жив ли данный человек
        private bool isLife;

        // заготовка функции "говорить"
        public void Talk()
        {
            // генерируем случайное число от 1 до 3
            int random_talk = rnd.Next(1, 4);

            // объявляем переменную, в которой мы будет хранить строку
            string tmp_str = "";

            switch (random_talk)
            {
                case 1: // если 1 - называем свое имя
                    { 
                        // генерируем текст сообщения
                        tmp_str = "Привет меня зовут" + Name + ", рад познакомится";
                        // завершаем оператор выбора
                        break;
                    }

                case 2: // возраст
                    {
                        // генерируем текст сообщения
                        tmp_str = "Мне" + Age + ". А тебе?";
                        // завершаем оператор выбора
                        break;
                    }
                case 3: // говорим о своем здоровье
                    {
                        // в зависимости от параметра здоровья
                        if (Health > 50)
                            tmp_str = "Да я здоров как бык!";
                        else
                            tmp_str = "Со здоровьем у меня хреново, дожить бы до" + (Age + 10).ToString();

                        // завершаем оператор выбора
                        break;
                    }
                    
            }
            // выводим в консоль сгенерированное сообщение
            System.Console.WriteLine(tmp_str);
        }

        // заготовка функции идти
        public void Go()
        {
            // если объект жив
            if (isLife == true)
            {
                // если показатель более 40
                // считаем объект здоровым
                if (Health > 40)
                {
                    // генерируем строку текста
                    string outString = Name + " мирно прогуливается по городу";
                    // выводим в консоль
                    System.Console.WriteLine(outString);
                }
                else
                {
                    // генерируем строку текста
                    string outString = Name + " болен и не может гулять по городу";
                    // выводим в консоль
                    System.Console.WriteLine(outString);
                }
            }
            else
            {
                // генерируем строку текста
                string outString = Name + " не может идти, он умер";
                System.Console.WriteLine(outString);
            }
        }

        public void Kill()
        {
            // устанавливаем значение isLife (жив)
            // в false
            isLife = false;
            System.Console.WriteLine(Name + " умер");
        }

        // функция, возвращаюшая показатель - жив ли данный человек.
        public bool IsAlive()
        {
            // возвращаем значение, к которому мы не можем
            // обратиться напрямую из вне класса
            // так как оно имеет статут private
            return isLife;
        }

    }
}
