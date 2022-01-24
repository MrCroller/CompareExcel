using Microsoft.Office.Interop.Excel;

namespace CompareExcel;

class ExcelUse : IDisposable
{
    private Application app;
    private Workbook workbook;
    private string? filePatch;

    public ExcelUse()
    {
        app = new Application();
    }

    internal bool Open(string filePatch)
    {
        try
        {
            workbook = app.Workbooks.Open(filePatch);
            this.filePatch = filePatch;


            return true;
        }
        catch (Exception ex)
        {
            Console.WriteLine(ex.Message);
            return false;
        }
    }

    internal void Save(string filePatch)
    {
        try
        {
            if (!string.IsNullOrEmpty(filePatch))
            {
                workbook.SaveAs(filePatch);
                this.filePatch = null;
            }
            else
            {
                workbook.Save();
            }
            
        }
        catch (Exception ex) { Console.WriteLine(ex.Message); }
    }

    public void Dispose()
    {
        try
        {
            workbook.Close();
        }
        catch (Exception ex) { Console.WriteLine(ex.Message); }
    }
}

