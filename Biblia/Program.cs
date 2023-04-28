using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;


namespace Biblia
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Functional fun = new Functional();
            int answer;
            while (true)
            {
                int k = 0;
                Console.WriteLine("\n1-Просмотр списка читателей, не вернувших книгу\n" +
                "2- Посмотреть перечень книг в библиотеке\n" +
                "3- Добавить новую запись\n" +
                "4- Отметить возврат книги\n" +
                "5- Добавить новую книгу в архив\n");
                try
                {
                    answer = Convert.ToInt32(Console.ReadLine());
                    switch (answer)
                    {
                        case 1:
                            fun.conclusionReaders();
                            break;
                        case 2:
                            fun.bookList();
                            break;
                        case 3:
                            fun.addingNewEntry();
                            break;
                        case 4:
                            fun.returnMarkBook();
                            break;
                        case 5:
                            fun.addingBook(); 
                            break;
                        default: 
                            k++; 
                            break;
                    }
                }
                catch
                {
                    Console.WriteLine("Введите корректное значение!");
                }
                if(k != 0)
                {
                    break;
                }
            }
            Console.WriteLine("Программа завершила свою работу!");
            Console.ReadLine();
        }
    }
}