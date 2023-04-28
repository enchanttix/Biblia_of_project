using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace Biblia
{
    internal class Functional
    {
        int rowCount;
        int colCount;
        Excel.Application excelApp = new Excel.Application();//создаем объект приложения
        string pathFile = Environment.CurrentDirectory + "\\list.xlsx";

        Workbook ExcelBook;
        public Functional()
        {

        }

        /// <summary>
        /// Открытие файла
        /// </summary>
        /// <param name="numberList">Номер листа</param>
        /// <returns>Данные</returns>
        Range fileAccess(int numberList)
        {
            ExcelBook = excelApp.Workbooks.Open(pathFile);
            _Worksheet worksheet = (Worksheet)ExcelBook.Worksheets[numberList];//выбираем лист
            Range excelRange = worksheet.UsedRange;//найти используемые ячейки в массиве
            return excelRange;
        }

        /// <summary>
        /// Вывод должников
        /// </summary>
        public void conclusionReaders()
        {
            Range excelRange = fileAccess(1);
            rowCount = excelRange.Rows.Count;
            colCount = excelRange.Columns.Count;

            for (int i = 2; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {                   
                    if ((excelRange.Cells[i, j] != null) && (excelRange.Cells[i, j].Value2 != null) && (excelRange.Cells[i, 6].Value2 == null))
                    { 
                        Console.Write(excelRange.Cells[i, j].Value2.ToString() + "\t");
                        if (j == (colCount - 1))
                        {
                            Console.Write("\r\n");
                        } 
                    }
                    
                }
            }      
            excelApp.Quit();
        }

        /// <summary>
        /// Вывод книг
        /// </summary>
        public void bookList()
        {
            Range excelRange = fileAccess(2);
            rowCount = excelRange.Rows.Count;
            colCount = excelRange.Columns.Count;
            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {
                    if (excelRange.Cells[i, j].Value2 == null)
                    {
                        Console.Write("\t");
                    }
                    if ((excelRange.Cells[i, j] != null) && (excelRange.Cells[i, j].Value2 != null))
                    {
                        Console.Write(excelRange.Cells[i, j].Value2.ToString() + "\t");                      
                    }
                }
             Console.WriteLine();
            }
            excelApp.Quit();
        }

        /// <summary>
        /// Запись о возврате
        /// </summary>
        public void returnMarkBook()
        {
            Range excelRange = fileAccess(1);
            rowCount = excelRange.Rows.Count;
            colCount = excelRange.Columns.Count;
            Console.WriteLine("Введите ФИО читателя");
            string reader = Console.ReadLine();

            for (int i = 1; i <= rowCount; i++)
            {
                for (int j = 1; j <= colCount; j++)
                {                   
                    if ((excelRange.Cells[i, j] != null) && (excelRange.Cells[i, j].Value2 != null) && (excelRange.Cells[i,2].Value2==reader) && (excelRange.Cells[i, 6].Value2 == null))
                    {
                        Console.Write(excelRange.Cells[i, j].Value2.ToString() + "\t");
                        if (j == (colCount - 1))
                        {
                            Console.Write("\r\n");
                            Console.WriteLine("Введите дату возврата (пример: 10 10 2010)");
                            excelRange.Cells[i, 6] = Console.ReadLine();
                            ExcelBook.Save();
                        }                       
                    }
                }
            }
            excelApp.Quit();
        }

        /// <summary>
        /// Новая запись о взятии книг
        /// </summary>
        public void addingNewEntry()
        {
            Range excelRange = fileAccess(1);
            rowCount = excelRange.Rows.Count;
            colCount = excelRange.Columns.Count;
            for (int i = 1; i < colCount; i++)
            {
                switch(i)
                {
                    case 1:                        
                            Console.WriteLine("Введите фамилию библиотекаря");
                            break;                       
                    case 2:                       
                            Console.WriteLine("Введите ФИО читателя");
                            break;                       
                    case 3:                        
                            Console.WriteLine("Введите название книги");
                            break ;                      
                    case 4:                       
                            Console.WriteLine("Введите автора");
                            break;                        
                    case 5:                        
                            Console.WriteLine("Введите дату взятия (пример: 10 10 2010)");
                            break;                                          
                        
                }
                excelRange.Cells[rowCount+1, i] = Console.ReadLine();
            }
            ExcelBook.Save();
            excelApp.Quit();
        }

        /// <summary>
        /// Новая запись  книг
        /// </summary>
        public void addingBook()
        {
            Range excelRange = fileAccess(2);
            rowCount = excelRange.Rows.Count;
            colCount = excelRange.Columns.Count;
            for (int i = 1; i < colCount+1; i++)
            {
                switch (i)
                {
                    case 1:
                        Console.WriteLine("Введите автора");
                        break;
                    case 2:
                        Console.WriteLine("Введите название произведения");
                        break;
                    case 3:
                        Console.WriteLine("Введите издательство");
                        break;
                    

                }
                excelRange.Cells[rowCount + 1, i] = Console.ReadLine();
            }
            ExcelBook.Save();
            excelApp.Quit();
        }
    }
}
