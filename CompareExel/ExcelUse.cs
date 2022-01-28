using System.Collections.Generic;
using Excel = Microsoft.Office.Interop.Excel;

namespace CompareExcel;

class ExcelUse : IDisposable
{
    public delegate void ErrorQuit();
    /// <summary>
    /// Событие на случай ошибки
    /// </summary>
    public event ErrorQuit? Turn_nextEV;

    private readonly Excel.Application app;
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
            Visible = false
        };
        if (app == null)
        {
            Turn_nextEV?.Invoke();
            throw new Exception("Приложение Excel не запущенно, убедитесь что пакет office установлен");
        }
    }

    /// <summary>
    /// Чтение файла из указанного пути (возвращает DataTable)
    /// </summary>
    /// <param name="filePatch"></param>
    /// <returns></returns>
    public System.Data.DataTable ReadFile(string filePatch)
    {
        try
        {
            Console.WriteLine($"\nReadFile {filePatch.Split(@"\")[^1]}");
            for (int i = 0; i < 30; i++)
                Console.Write("###");
            Console.WriteLine("\nSTART READ");

            System.Data.DataTable DT = new();

            this.Workbook = app.Workbooks.Open(filePatch);
            this.filePatch = filePatch;

            Sheet = (Excel.Worksheet)Workbook.Worksheets[1];
            if (Sheet == null)
            {
                Turn_nextEV?.Invoke();
                throw new Exception("Листа не существует, убедитесь в правивильности имени");
            }
            //Sheet.Activate();

            #region Приведения к Excel.Range возможно пригодится на более ранних версиях языка
            //Console.WriteLine((Sheet.UsedRange.Cells[1, 1] as Excel.Range).Text); 
            #endregion

            var useRange = Sheet.UsedRange;
            int rowsCount = useRange.Rows.Count;
            int columnsCount = useRange.Columns.Count;

            Console.WriteLine($"rowsCount: {rowsCount}\ncolumnsCount: {columnsCount}");

            // Именование колонок
            for (int i = 1; i <= columnsCount; i++)
            {
                dynamic columnName = useRange.Cells[1, i];
                DT.Columns.Add(columnName.Text, columnName.GetType());
            }
            // Получение строк
            List<object> listRow = new(); // Лист строки
            for (int i = 2; i <= rowsCount; i++)
            {
                //Console.WriteLine($"new row [{i-1}]");
                listRow.Clear();
                for (int j = 1; j <= columnsCount; j++)
                {
                    //Console.WriteLine($"[{i-1},{j}]Text: {useRange.Cells[i, j].Text}");
                    listRow.Add(useRange.Cells[i, j]); // Заполнение листа строки
                }
                DT.Rows.Add(listRow.ToArray()); // Добавление новой строки в DT
            }

            Console.WriteLine("FINISH READ");
            return DT;
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка считывания: {ex}");
            this.filePatch = null;

            return null;
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
            Console.WriteLine($"Convert to DataTable ");
            this.Sheet.Name = SheetName; // Именование таблицы формирования
            var useRange = Sheet.UsedRange;

            // Именование колонок
            char _excelHeader = 'A';
            foreach (System.Data.DataColumn column in DT.Columns)
            {
                Sheet.Cells[1, _excelHeader.ToString()] = column.ColumnName;
                _excelHeader++;
            }
            // Остальная информация
            for (int i = 0; i < DT.Rows.Count; i++)
            {
                _excelHeader = 'A';
                int arrLength = DT.Rows[i].ItemArray.Length;
                for (int j = 0; j < arrLength; j++)
                {
                    Sheet.Cells[i + 2, _excelHeader] = DT.Rows[i].ItemArray[j];
                }
                _excelHeader++;
            }
        }
        catch (Exception ex)
        {
            Console.WriteLine($"Ошибка конвертирования: {ex.Message}");
        }
    }

    /// <summary>
    /// Сохранение файла
    /// </summary>
    public void Save()
    {
        try
        {
            Workbook.SaveAs(filePatch);
        }
        catch (Exception ex) { Console.WriteLine(ex.Message); }
    }

    /// <summary>
    /// Сохранение файла в указанный путь
    /// </summary>
    /// <param name="fileFolder">Путь к папке</param>
    /// <param name="fileName">Имя файла</param>
    public void SaveAs(string fileFolder, string fileName)
    {
        filePatch = $@"{fileFolder}/{fileName}";
        Save();
    }

    /// <summary>
    /// Освобождение ресурсов, закрытие рабочей книги
    /// </summary>
    public void Dispose()
    {
        try
        {
            Workbook.Close();
            app.Quit();
            Console.WriteLine("\nDispose. Очистка процессов");
        }
        catch (Exception ex) { Console.WriteLine(ex.Message); }
    }

    /// <summary>
    /// Закрытие приложения (ручной запуск)
    /// </summary>
    protected internal void Quit() => app.Quit();
}

