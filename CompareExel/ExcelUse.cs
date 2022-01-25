using Excel = Microsoft.Office.Interop.Excel;

namespace CompareExcel;

class ExcelUse : IDisposable
{
    private readonly Excel.Application app;
    /// <summary>
    /// Рабочая книга
    /// </summary>
    public Excel.Workbook Workbook;
    /// <summary>
    /// Страница книги Excel
    /// </summary>
    private Excel.Worksheet Sheet;
    private readonly string SheetName = "Лист1";

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
    }

    /// <summary>
    /// Инициализация приложения
    /// </summary>
    /// <param name="SheetName">Имя листа</param>
    public ExcelUse(string SheetName)
    {
        app = new Excel.Application();
        this.SheetName = SheetName;
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

            this.Sheet = Workbook.Worksheets[SheetName];

            // TODO: Создать нормальное формирование таблицы как в методе Convert
            foreach (System.Data.DataRow row in Sheet.Rows)
            {
                DT.Rows.Add(row);
            }

            return DT;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
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
            Console.WriteLine(ex.Message);
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

