using System.Data;
using System.Diagnostics;
using System.Linq;

namespace CompareExcel;

public class Program
{
    /// <summary>
    /// Директория извлекаемых таблиц
    /// </summary>
    static private string directory;
    /// <summary>
    /// Директория вывода файлов
    /// </summary>
    static private string dirOut;
    /// <summary>
    /// Название консоли
    /// </summary>
    static private string title = "Excel Merge";

    /// <summary>
    /// Вывод информации в консоль
    /// </summary>
    static INFOConsole INFO = INFOConsole.Main;
    /// <summary>
    /// Формат даты месяц и год
    /// </summary>
    struct DateMY
    {
        public int month;
        public int year;

        public string GetDate() => $"{month}.{year}";
    }

    static void Main(string[] args)
    {
        AnalizArgs(args);

        Stopwatch timer = new();
        Console.Title = title;
        Console.CursorVisible = false;

        Console.Clear();
        timer.Start(); // Таймер

        Dictionary<DateMY, DataTable> dictDT = new();
        List<DataTable> allDT = ReadRange(directory);
        foreach (DataTable dt in allDT)
        {
            string my_str = dt.Rows[0]["ns1:DocDate"].ToString();
            DateMY date = new()
            {
                month = Convert.ToInt32(my_str.ToString().Substring(3, 2)),
                year = Convert.ToInt32(my_str.ToString().Substring(6, 4))
            };

            if (dictDT.ContainsKey(date))
            {
                // Соединение таблиц. Сохранение данных и добавление других колонок
                dictDT[date].Merge(dt, true, MissingSchemaAction.Add);
            }
            else
            {
                dictDT.Add(date, dt);
            }
        }

        int counter = 1; // Счетчик для прогресса
        foreach (KeyValuePair<DateMY, DataTable> dt in dictDT)
        {
            Console.Title = $"DataTable: {dt.Key.GetDate()} | Прогресс = {Math.Round((double)(counter * 100 / dictDT.Count))}% [{counter}/{dictDT.Count}]";

            if ((int)INFO > 0) Console.WriteLine($"\nУдаление дублей. Ключ таблицы: {dt.Key.GetDate()}");
            dt.Value.AsEnumerable().Distinct(DataRowComparer.Default);

            using ExcelUse excel = new(INFO);
            excel.Convert(dt.Value);
            excel.SaveAs(dirOut, GetFileName(directory, dt.Key.month, dt.Key.year));
            counter++;
        }

        timer.Stop();

        // TODO: Сделать правильный, корректный, вывод затраченного времени
        var time = timer.ElapsedMilliseconds;
        int time_s = (int)time / 1000;
        int time_m = (int)time / 1000 / 60;

        Console.WriteLine($"Потрачено времени: {time_m}m{time_s - (time_m * 60)}s");
    }

    #region Параметры запуска

    /// <summary>
    /// Анализ входящих аргументов
    /// </summary>
    /// <param name="args">Массив аргументов</param>
    private static void AnalizArgs(string[] args)
    {
        Console.WriteLine($"args.Length: {args.Length}");
        // Анализ аргументов
        for (int i = 0; i < args.Length; i++)
        {
            Console.WriteLine($"i: {i}");
            switch (args[i])
            {
                //TODO: Ошибка с параметрами запуска при отсутствии последнего
                case "-help":
                    Console.WriteLine("-----Commands-----\n-help (показать команды)\n-p (Папка для чтения)\n-o (Папка для вывода объединенного файла)\n-i (Уровень информирования, выбор из 3 доступных: None - Без консольного вывода, Main - главные процессы [по умолчанию], All - Полный вывод, аж на каждую ячеечку. Много букав!)\n------------------");
                    break;
                case "-p":
                    string patchDir = PatchSplit(args, ref i);
                    if (Directory.Exists(patchDir))
                    {
                        Console.WriteLine($"\n{patchDir}\n"); //to_delete
                        directory = patchDir;
                    }
                    else
                    {
                        Console.WriteLine("Не задана папка для чтения или ее не существует");
                        Console.WriteLine($"\n{patchDir}\n"); //to_delete
                    }
                    break;
                case "-o":
                    string patchOut = PatchSplit(args, ref i);
                    if (Directory.Exists(patchOut))
                    {
                        Console.WriteLine($"\n{patchOut}\n"); //to_delete
                        dirOut = patchOut;
                    }
                    else
                    {
                        Console.WriteLine("Не задана папка для вывода или ее не существует");
                        Console.WriteLine($"\n{patchOut}\n"); //to_delete
                    }
                    break;
                case "-i":
                    switch (args[i + 1].ToLower())
                    {
                        case "none": INFO = INFOConsole.None; break;
                        case "main": INFO = INFOConsole.Main; break;
                        case "all": INFO = INFOConsole.All; break;
                        default:
                            Console.WriteLine("Вариант не распознан. Вывод по умолчанию: Main");
                            break;
                    }
                    break;
            }
        }
        if (directory == null || dirOut == null)
        {
            throw new Exception("Необходимо ввести путь к папке для сортировки и выходную папку для объединненных файлов");
        }
        Console.ReadKey();
    }

    /// <summary>
    /// Объединяет строку пути если в ней есть пробелы
    /// </summary>
    /// <param name="args">Массив аргументов</param>
    /// <param name="i">Текущий индекс элемента</param>
    /// <returns></returns>
    private static string PatchSplit(string[] args, ref int i)
    {
        string? strDir = null;
        // Если в пути есть пробелы объединяет строки
        for (int j = i; !args[j].StartsWith('-'); j++)
        {
            strDir += args[j + 1] + " ";
            i++;
        }
        return strDir.Trim();
    }

    #endregion

    /// <summary>
    /// Метод читает все файлы в каталоге и возвращает лист таблиц DataTable
    /// </summary>
    /// <param name="dirName">Путь к папке</param>
    /// <returns></returns>
    private static List<DataTable> ReadRange(string dirName)
    {
        try
        {
            string[] filesDir = Directory.GetFiles(dirName, "*.xlsx"); // Сканирует файлы формата Excel
            string folder = dirName.Split(@"\")[^1];

            List<DataTable> list = new();

            Console.WriteLine($"Сканирование директории: {dirName}");

            using ExcelUse excelApp = new(INFO);
            for (int i = 0; i < filesDir.Length; i++)
            {
                Console.Title = $"Папка: {folder} | Прогресс = {Math.Round((double)((i + 1) * 100 / filesDir.Length))}% [{i + 1}/{filesDir.Length}]";
                if ((int)INFO > 0) Console.WriteLine($"\n[{i + 1}/{filesDir.Length}]");
                list.Add(excelApp.ReadFile(filesDir[i]));
            }
            Console.Title = title;

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
        // Если месяц меньше 10 добавляется 0, для формирования
        string mounthSTR = (mounth > 10) ? mounth.ToString() : "0" + mounth.ToString();
        // Если путь с названием файла, то берется вторая подстрочка с конца (имя папки)
        return $@"{DirFolder.Split(@"\")[^(DirFolder.EndsWith(".xlsx") ? 2 : 1)]}_{mounthSTR}_{year}";
    }
}