using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Text.RegularExpressions;

namespace CompareExcel;

public class Program
{
    /// <summary>
    /// Директория извлекаемых таблиц
    /// </summary>
    static public string directory = @"C:\Users\pvslavinsky\Desktop\ФКР\Результаты\Постановление о запрете на регистрационные действия в отношении транспортных средств";

    static void Main(string[] args)
    {
        Stopwatch timer = new();
        timer.Start(); // Таймер

        string[] filesDir = Directory.GetFiles(directory, "*.xlsx"); // Сканирует файлы формата Excel

        using (ExcelUse ex = new())
        {
            DataTable dt = ex.ReadFile($@"{directory}\piev_65011250350281.xlsx");
            //DataTable dt = new DataTable() { };
            Console.WriteLine($"NameDT: {dt.TableName}, Rows.Count:{dt.Rows.Count}");
            ex.Convert(dt);
            ex.SaveAs(@"C:\Users\pvslavinsky\Desktop\ФКР\Результаты\Full", GetFileName(directory));
        }

        timer.Stop();
        Console.WriteLine($"Директория: {directory}\n Потрачено времени: {timer.ElapsedMilliseconds} ms.");
    }

    /// <summary>
    /// Метод читает все файлы в каталоге и возвращает лист таблиц DataTable
    /// </summary>
    /// <param name="dirName">Путь к папке</param>
    /// <returns></returns>
    static List<DataTable> ReadRange(string dirName)
    {
        string[] filesDir = Directory.GetFiles(dirName, "*.xlsx"); // Сканирует файлы формата Excel

        foreach (string filePatch in filesDir)
        {
            try
            {
                using ExcelUse excelApp = new();
                excelApp.ReadFile(filePatch);
            }
            catch (Exception ex) { Console.WriteLine(ex.Message); }
        }

        return null;
    }

    /// <summary>
    /// Создание имени файла
    /// </summary>
    /// <param name="DirFolder">Директория извлекаемых таблиц</param>
    /// <returns></returns>
    private static string GetFileName(string DirFolder) => GetFileName(DirFolder, DateTime.Now.Month, DateTime.Now.Year);

    /// <summary>
    /// Создание имени файла
    /// </summary>
    /// <param name="DirFolder">Директория извлекаемых таблиц</param>
    /// <param name="mounth">Месяц</param>
    /// <param name="year">Год</param>
    /// <returns></returns>
    private static string GetFileName(string DirFolder, int mounth, int year)
    {
        // Если путь с названием файла, то берется вторая подстрочка с конца (имя папки)
        return $"{DirFolder.Split(@"\")[^(DirFolder.EndsWith(".xlsx") ? 2 : 1)]}_{mounth}_{year}";
    }
}