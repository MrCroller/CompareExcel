global using System;
global using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace CompareExcel;

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

    public ExcelUse()
    {
        app = new Excel.Application
        {
            Visible = false,
            SheetsInNewWorkbook = 1
        };
        if (app == null)
        {
            throw new Exception("Приложение Excel не запущенно, убедитесь что пакет office установлен");
        }
        CConsole.Print("\nNew Application", col: ConsoleColor.DarkMagenta);
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
            // TODO: Ввести дополнительные параметры для открытия во избежании бага
            this.Workbook = app.Workbooks.Open(filePatch);
            this.filePatch = filePatch;

            System.Data.DataTable DT = new();

            CConsole.Print("START READ File: ", col: ConsoleColor.DarkRed, newLine: false);
            CConsole.Print(filePatch.Split(@"\")[^1]);

            Sheet = (Excel.Worksheet)Workbook.Worksheets[1];
            if (Sheet == null)
            {
                throw new Exception("Листа не существует, убедитесь в правивильности имени");
            }
            Sheet.Activate();

            var useRange = Sheet.UsedRange;
            int rowsCount = useRange.Rows.Count;
            int columnsCount = useRange.Columns.Count;

            CConsole.Print($"rowsCount: {rowsCount}\ncolumnsCount: {columnsCount}");
            CConsole.Separator(strInfo: INFOConsole.All);

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
                CConsole.Print($"new row [{i - 1}]\n", strInfo: INFOConsole.All);
                listRow.Clear();
                for (int j = 1; j <= columnsCount; j++)
                {
                    string cell = useRange.Cells[i, j].Text;
                    CConsole.Print($"[{i - 1},{j}]Text: {cell}", strInfo: INFOConsole.All);
                    listRow.Add(cell); // Заполнение листа строки
                }
                CConsole.SetLine(CConsole.ProgressBar(i - 1, rowsCount, "Прогресс чтения: ", spiner: true));
                DT.Rows.Add(listRow.ToArray()); // Добавление новой строки в DT
                CConsole.WriteLine(INFOConsole.All);
            }
            CConsole.Separator(strInfo: INFOConsole.All);
            CConsole.WriteLine(INFOConsole.Main);
            CConsole.Print("FINISH READ", col: ConsoleColor.Green);

            return DT;
        }
        catch (Exception ex)
        {
            CConsole.Print($"Ошибка считывания: {ex}", INFOConsole.None, ConsoleColor.Red);
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

            CConsole.Print($"START CONVERT to DataTable ", col: ConsoleColor.Red);

            // Именование колонок
            for (int i = 1; i < DT.Columns.Count; i++)
            {
                Sheet.Cells[1, i] = DT.Columns[i - 1].ColumnName;
            }

            // Остальная информация
            int maxRows = DT.Rows.Count;
            for (int i = 0; i < maxRows; i++)
            {
                //CConsole.SetLine(CConsole.Progress(i, maxRows, "Прогресс конвертации: ", "rows: ", true));
                CConsole.SetLine(CConsole.ProgressBar(i, maxRows, "Прогресс конвертации: ", spiner: true));
                CConsole.Print($"new row [{i}]\n", strInfo: INFOConsole.All);

                for (int j = 0; j < DT.Columns.Count; j++)
                {
                    object? cell = DT.Rows[i].ItemArray[j].ToString();
                    CConsole.Print($"[{i + 1},{j + 1}]Text: {cell}", strInfo: INFOConsole.All);
                    Sheet.Cells[i + 2, j + 1] = cell;
                }
                CConsole.WriteLine(INFOConsole.All);
            }
            CConsole.WriteLine();
            CConsole.Print("FINISH CONVERT", col: ConsoleColor.Green);
        }
        catch (Exception ex)
        {
            CConsole.Print($"Ошибка конвертирования: {ex.Message}", INFOConsole.None, ConsoleColor.Red);
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
                CConsole.Print($"Файл {filePatch} уже есть в директории. Удаление", INFOConsole.None, ConsoleColor.DarkRed);
                File.Delete(filePatch);
            }
            Workbook.SaveAs(filePatch, Excel.XlFileFormat.xlWorkbookDefault, Type.Missing, Type.Missing, false, false, Excel.XlSaveAsAccessMode.xlNoChange, Excel.XlSaveConflictResolution.xlLocalSessionChanges, false);
        }
        catch (Exception ex)
        {
            CConsole.Print($"Ошибка сохранения: {ex.Message}", col: ConsoleColor.Red);
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
        CConsole.Print("Dispose. Очистка процессов", col: ConsoleColor.Yellow);

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
        CConsole.Print("CloseWB", col: ConsoleColor.Yellow);
        Workbook.Close(false);
    }
}

