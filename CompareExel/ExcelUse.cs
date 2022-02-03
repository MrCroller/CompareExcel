global using System;
global using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace CompareExcel;

/// <summary>
/// Кол-во данных для вывода в консоль
/// </summary>
enum INFOConsole
{
    /// <summary>
    /// Без вывода данных
    /// </summary>
    None,
    /// <summary>
    /// Вывод информации о главных процессах
    /// </summary>
    Main,
    /// <summary>
    /// Полный вывод данных
    /// </summary>
    All
}

class ExcelUse : IDisposable
{
    private Excel.Application? app;
    /// <summary>
    /// Рабочая книга
    /// </summary>
    public Excel.Workbook Workbook;
    /// <summary>
    /// Страница книги Excel
    /// </summary>
    private Excel.Worksheet Sheet;
    /// <summary>
    /// Имя листа в итоговом файле
    /// </summary>
    private readonly string SheetName = "Summ";

    /// <summary>
    /// Путь к файлу
    /// </summary>
    internal string? filePatch;
    /// <summary>
    /// Уровень информирования (данных в консоль)
    /// </summary>
    private INFOConsole INFO;

    public ExcelUse(INFOConsole info)
    {
        this.INFO = info;
        app = new Excel.Application
        {
            Visible = false,
            SheetsInNewWorkbook = 1
        };
        if (app == null)
        {
            throw new Exception("Приложение Excel не запущенно, убедитесь что пакет office установлен");
        }
        if ((int)INFO > 0) CColor("\nNew Application", ConsoleColor.DarkMagenta);
    }

    /// <summary>
    /// Чтение файла из указанного пути (возвращает DataTable)
    /// </summary>
    /// <param name="filePatch">Путь к папке</param>
    /// <param name="info">Нужно ли выводить данные в консоль?</param>
    /// <returns></returns>
    public System.Data.DataTable ReadFile(string filePatch)
    {
        try
        {
            this.Workbook = app.Workbooks.Open(filePatch);
            this.filePatch = filePatch;

            System.Data.DataTable DT = new();

            if ((int)INFO > 0)
            {
                CColor("START READ File: ", ConsoleColor.DarkRed, false);
                Console.WriteLine(filePatch.Split(@"\")[^1]);
            }

            Sheet = (Excel.Worksheet)Workbook.Worksheets[1];
            if (Sheet == null)
            {
                throw new Exception("Листа не существует, убедитесь в правивильности имени");
            }
            Sheet.Activate();

            var useRange = Sheet.UsedRange;
            int rowsCount = useRange.Rows.Count;
            int columnsCount = useRange.Columns.Count;

            if ((int)INFO > 0) Console.WriteLine($"rowsCount: {rowsCount}\ncolumnsCount: {columnsCount}");
            if ((int)INFO > 1)
            {
                Console.WriteLine();
                CSeparator();
            }
            // Именование колонок
            for (int i = 1; i <= columnsCount; i++)
            {
                dynamic columnName = useRange.Cells[1, i];
                DT.Columns.Add(columnName.Text);
            }
            // Получение строк
            List<object> listRow = new(); // Лист строки
            for (int i = 2; i <= rowsCount; i++)
            {
                if ((int)INFO > 1) Console.WriteLine($"new row [{i - 1}]\n");
                listRow.Clear();
                for (int j = 1; j <= columnsCount; j++)
                {
                    string cell = useRange.Cells[i, j].Text;
                    if ((int)INFO > 1) Console.WriteLine($"[{i - 1},{j}]Text: {cell}");
                    listRow.Add(cell); // Заполнение листа строки
                }
                DT.Rows.Add(listRow.ToArray()); // Добавление новой строки в DT
                if ((int)INFO > 1) Console.WriteLine();
            }
            if ((int)INFO > 1)
            {
                CSeparator();
                Console.WriteLine();
            }
            if ((int)INFO > 0) CColor("FINISH READ", ConsoleColor.Green);

            return DT;
        }
        catch (Exception ex)
        {
            CColor($"Ошибка считывания: {ex}", ConsoleColor.Red);
            this.filePatch = null;

            return null;
        }
        finally
        {
            CloseWB();
        }
    }

    /// <summary>
    /// Преобразование DataTable в рабочую страницу
    /// </summary>
    /// <param name="DT"></param>
    public void Convert(System.Data.DataTable DT)
    {
        try
        {
            this.Workbook = app.Workbooks.Add(Type.Missing);
            this.Sheet = (Excel.Worksheet)app.Worksheets[1];
            this.Sheet.Name = SheetName; // Именование таблицы формирования

            if ((int)INFO > 0) CColor($"START CONVERT to DataTable ", ConsoleColor.Red);

            // Именование колонок
            for (int i = 1; i < DT.Columns.Count; i++)
            {
                Sheet.Cells[1, i] = DT.Columns[i - 1].ColumnName;
            }

            // Остальная информация
            int maxRows = DT.Rows.Count;
            for (int i = 0; i < maxRows; i++)
            {
                int pr_length = 0;
                if ((int)INFO == 1) CSetLine($"Прогресс конвертации: {Math.Round((double)((i + 1) * 100 / maxRows))}% rows[{i + 1}/{maxRows}]"); // Прогресс при уровне информирования Main
                if ((int)INFO > 1) Console.WriteLine($"new row [{i}]\n");
                for (int j = 0; j < DT.Columns.Count; j++)
                {
                    object? cell = DT.Rows[i].ItemArray[j].ToString();
                    if ((int)INFO > 1) Console.WriteLine($"[{i + 1},{j + 1}]Text: {cell}");
                    Sheet.Cells[i + 2, j + 1] = cell;
                }
                if ((int)INFO > 1) Console.WriteLine();
            }
            if ((int)INFO == 1) Console.WriteLine();
            if ((int)INFO > 0) CColor("FINISH CONVERT", ConsoleColor.Green);
        }
        catch (Exception ex)
        {
            CColor($"Ошибка конвертирования: {ex.Message}", ConsoleColor.Red);
        }
    }

    /// <summary>
    /// Сохранение файла
    /// </summary>
    public void Save()
    {
        try
        {
            // Сохранение с перезаписью
            if (File.Exists(filePatch))
            {
                CColor($"Файл {filePatch} уже есть в директории. Удаление", ConsoleColor.DarkRed);
                File.Delete(filePatch);
            }
            Workbook.SaveAs(filePatch, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, false);
        }
        catch (Exception ex)
        {
            CColor($"Ошибка сохранения: {ex.Message}", ConsoleColor.Red);
        }
    }
    /// <summary>
    /// Сохранение файла в указанный путь
    /// </summary>
    /// <param name="fileFolder">Путь к папке</param>
    /// <param name="fileName">Имя файла</param>
    public void SaveAs(string fileFolder, string fileName)
    {
        filePatch = $@"{fileFolder}\{fileName}.xlsx";
        Save();
    }

    /// <summary>
    /// Освобождение ресурсов, закрытие рабочей книги
    /// </summary>
    public void Dispose()
    {
        if ((int)INFO > 0) CColor("Dispose. Очистка процессов", ConsoleColor.Yellow);

        app.Quit();
        //while(Marshal.ReleaseComObject(app) != 0) { }
        app = null;
        GC.Collect();
        //GC.WaitForPendingFinalizers();
    }

    /// <summary>
    /// Закрытие приложения (ручной запуск)
    /// </summary>
    protected internal void Quit() => app.Quit();
    /// <summary>
    /// Закрытие рабочей книги (ручной запуск)
    /// </summary>
    protected internal void CloseWB()
    {
        if ((int)INFO > 0) CColor("CloseWB", ConsoleColor.Yellow);
        Workbook.Close(false);
    }

    #region Консоль
    /// <summary>
    /// Пишет текст указанным цветом
    /// </summary>
    /// <param name="str">Текст для вывода в консоль</param>
    /// <param name="col">Цвет которым будет написан текст</param>
    /// <param name="nonNewLine">Добавлять новую строку (WriteLine)?</param>
    private static void CColor(string str, ConsoleColor col = ConsoleColor.White, bool nonNewLine = true)
    {
        ConsoleColor def = Console.ForegroundColor; // Получение стандартного цвета

        Console.ForegroundColor = col;

        if (nonNewLine) Console.WriteLine(str);
        else Console.Write(str);

        Console.ForegroundColor = def; // Возвращение стандартного цвета
    }
    /// <summary>
    /// Разделитель информации в строке
    /// </summary>
    /// <param name="lenght">Длина деления</param>
    private static void CSeparator(int lenght)
    {
        for (int i = 0; i < lenght; i++)
            Console.Write("#");
        Console.WriteLine();
    }
    /// <summary>
    /// Разделитель информации в строке
    /// </summary>
    /// <param name="min">Минимальное кол-во символов</param>
    /// <param name="max">Максимальное кол-во символов</param>
    private static void CSeparator(int min = 30, int max = 50)
    {
        Random rnd = new();
        int ss = rnd.Next(min, max);
        CSeparator(ss);
    }
    /// <summary>
    /// Пишет в консоль с той же строчки (перезаписывает)
    /// </summary>
    /// <param name="str">Строка</param>
    /// <param name="horiaontal">Горизонтальное смещение</param>
    private static void CSetLine(string str, int horiaontal = 0)
    {
        Console.SetCursorPosition(horiaontal, Console.GetCursorPosition().Top);
        Console.Write(str);
    }
    #endregion
}

