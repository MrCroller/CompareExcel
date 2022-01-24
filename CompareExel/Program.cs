using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using ExcelDataReader;

namespace CompareExcel;

class Program
{
    static void Main(string[] args)
    {
        Stopwatch sw_total = new Stopwatch();
        sw_total.Start(); // Таймер

        sw_total.Stop();
    }

    /// <summary>
    /// Метод читает все файлы в каталоге и возвращает лист таблиц DataTable
    /// </summary>
    /// <param name="dirName">Путь к папке</param>
    /// <returns></returns>
    private List<DataTable> ReadRange(string dirName)
    {
        string[] filesDir = Directory.GetFiles(dirName, "*.xlsx"); // Сканирует файлы формата Excel

        foreach (string filePatch in filesDir)
        {
            try
            {
                using (var excelApp = new ExcelUse())
                {
                    excelApp.Open(filePatch);
                }
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        return null;
    }
}