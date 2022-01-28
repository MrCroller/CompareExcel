using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Linq;
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

        Dictionary<int, DataTable> dictDT = new();
        List<DataTable> allDT = ReadRange(directory);
        foreach (DataTable dt in allDT)
        {
            // Взятие со строки даты документа {DocDate} месяц и год формат: {042022}
            int month = Convert.ToInt32(dt.Rows[2]["ns1:DocDate"].ToString().Substring(2).Replace(".", ""));

            if (dictDT.Keys.Contains(month))
            {
                dictDT[month].Merge(dt);
            }
            else
            {
                dictDT.Add(month, dt);
            }
            // TODO: Напиши конвертирования по словарю, и сохранение
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
        try
        {
            List<DataTable> list = new();
            Console.WriteLine($"Сканирование директории: {dirName.Substring(0, dirName.LastIndexOf(@"\"))}");
            string[] filesDir = Directory.GetFiles(dirName, "*.xlsx"); // Сканирует файлы формата Excel
            foreach (string filePatch in filesDir)
            {
                using ExcelUse excelApp = new();
                list.Add(excelApp.ReadFile(filePatch));
            }
            return list;
        }
        catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
    }

    /// <summary>
    /// Создание имени файла
    /// </summary>
    /// <param name="DirFolder">Директория извлекаемых таблиц</param>
    /// <returns></returns>
    private static string GetFileName(string DirFolder) => GetFileName(DirFolder, DateTime.Now.Month, DateTime.Now.Year);

    /// <summary>
    /// Создание имени файла с установкой месяца и года
    /// </summary>
    /// <param name="DirFolder">Директория извлекаемых таблиц</param>
    /// <param name="mounth">Месяц</param>
    /// <param name="year">Год</param>
    /// <returns></returns>
    internal static string GetFileName(string DirFolder, int mounth, int year)
    {
        // Если путь с названием файла, то берется вторая подстрочка с конца (имя папки)
        return $"{DirFolder.Split(@"\")[^(DirFolder.EndsWith(".xlsx") ? 2 : 1)]}_{mounth}_{year}";
    }
}