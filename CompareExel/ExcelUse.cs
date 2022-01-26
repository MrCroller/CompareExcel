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
    public System.Data.DataTable? ReadFile(string filePatch)
    {
        try
        {
            System.Data.DataTable DT = new();

            this.Workbook = app.Workbooks.Open(filePatch);
            this.filePatch = filePatch;

            this.Sheet = (Excel.Worksheet)Workbook.Worksheets[0];
            if (this.Sheet == null)
            {
                Turn_nextEV?.Invoke();
                throw new Exception("Листа не существует, убедитесь в правивильности имени");
            }

            var rangeExcel = Sheet.UsedRange;
            Console.WriteLine((rangeExcel.Cells[2, 'J'] as Excel.Range).Text.ToString());
            Console.WriteLine(Sheet.Cells[2, 2]);

            // Именование колонок
            foreach (Excel.Range col in Sheet.Columns)
            {
                Console.WriteLine($"Col.name: {col[1]}");
                DT.Columns.Add(col.Name, col.GetType());
            }
            /*
            // Остальная информация
            foreach(Excel.Range row in Sheet.Rows)
            {
                Console.WriteLine($"row.Name = {row.Name}\nrow.count = {(row as System.Data.DataRow).ItemArray.Length}");
                DT.Rows.Add(row as System.Data.DataRow);
            }
            */
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
    /// Преобразование DataTable в рабочую страницу
    /// </summary>
    /// <param name="DT"></param>
    public void Convert(System.Data.DataTable DT)
    {
        try
        {
            this.Sheet.Name = "SummEx";
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
                    Sheet.Cells[i + 2, j] = DT.Rows[i].ItemArray[j];
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
    /// Освобождение ресурсов, закрытие рабочей книги
    /// </summary>
    public void Dispose()
    {
        try
        {
            Workbook.Close();
            app.Quit();
        }
        catch (Exception ex) { Console.WriteLine(ex.Message); }
    }

    /// <summary>
    /// Закрытие приложения (ручной запуск)
    /// </summary>
    protected internal void Quit() => app.Quit();
}

