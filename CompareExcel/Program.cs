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
    /// Указывает искать ли таблицы во вложенных папках
    /// </summary>
    static private bool main_Dir = false;
    /// <summary>
    /// Директория вывода файлов
    /// </summary>
    static private string dirOut;
    /// <summary>
    /// Название консоли
    /// </summary>
    static private string title = "Excel Merge ";

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

        if (main_Dir)
        {
            string[]? dirArr = Directory.GetDirectories(directory);
            if((int)INFO > 1) Console.WriteLine(string.Join("\n", dirArr));
            if (dirArr.Length == 0)
            {
                throw new Exception("Указанная папка не имеет вложенных папок");
            }

            for (int i = 0; i < dirArr.Length; i++)
            {
                Console.Title = $"dir:[{i}/{dirArr.Length}] ";

                List<DataTable> rr = ReadRange(dirArr[i]);
                Dictionary<DateMY, DataTable> dictDT = SortDT(rr);
                ConvAndSave(dictDT);
            }
        }
        else
        {
            List<DataTable> rr = ReadRange(directory);
            Dictionary<DateMY, DataTable> dictDT = SortDT(rr);
            ConvAndSave(dictDT);
        }

        timer.Stop();
        var time = timer.ElapsedMilliseconds;
        int time_s = (int)time / 1000;
        int time_m = (int)time / 1000 / 60;
        Console.WriteLine($"Потрачено времени: {time_m}m {time_s - (time_m * 60)}s");
    }

    /// <summary>
    /// Анализ входящих аргументов
    /// </summary>
    /// <param name="args">Массив аргументов</param>
    private static void AnalizArgs(string[] args)
    {
        // Анализ аргументов
        for (int i = 0; i < args.Length; i++)
        {
            switch (args[i])
            {
                case "-help":
                    Console.WriteLine("-----Commands-----\n-help (показать команды)\n-d (Папка для чтения)\n-n (Учитывать вложенные папки): true - учитывать, false - не учитывать [по умолчанию]\n-o (Папка для вывода объединенного файла)\n-i (Уровень информирования, выбор из 3 доступных: None - Без консольного вывода, Main - главные процессы [по умолчанию], All - Полный вывод, аж на каждую ячеечку. Много букав!)\nДля оболочки PowerShell оборачивайте путь в ковычки 'directory'\n------------------");
                    break;
                case "-d": // Папка для чтения
                    if (Directory.Exists(args[i + 1]))
                    {
                        directory = args[i + 1];
                    }
                    else
                    {
                        Console.WriteLine("Не задана папка для чтения или ее не существует");
                    }
                    break;
                case "-n": // Учитывать ли вложенные папки?
                    if (args[i + 1].ToLower() == "true")
                    {
                        main_Dir = true;
                    }
                    else if (args[i + 1].ToLower() == "false")
                    {
                        main_Dir = false;
                    }
                    else
                    {
                        Console.WriteLine("Нераспознан параметр вложенных папок -d");
                    }
                    break;
                case "-o": // Директория выхода
                    if (Directory.Exists(args[i + 1]))
                    {
                        dirOut = args[i + 1];
                    }
                    else
                    {
                        Console.WriteLine("Не задана папка для вывода или ее не существует");
                    }
                    break;
                case "-i": // Уровень информирования
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
    }

    #region Работа с файлами и Excel

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
            string tt = Console.Title;

            List<DataTable> list = new();

            Console.WriteLine($"Сканирование директории: {dirName}");

            using ExcelUse excelApp = new(INFO);
            for (int i = 0; i < filesDir.Length; i++)
            {
                Console.Title = tt + $"Папка: {folder} | Прогресс = {Math.Round((double)((i + 1) * 100 / filesDir.Length))}% [{i + 1}/{filesDir.Length}]";

                if ((int)INFO > 0) Console.WriteLine($"\n[{i + 1}/{filesDir.Length}]");
                list.Add(excelApp.ReadFile(filesDir[i]));
            }
            Console.Title = title;

            return list;
        }
        catch (Exception ex) { Console.WriteLine(ex.Message); return null; }
    }

    /// <summary>
    /// Сортирует словарь таблиц по месяцам
    /// </summary>
    /// <param name="allDT">Коллекция для сортировки</param>
    /// <returns></returns>
    private static Dictionary<DateMY, DataTable> SortDT(List<DataTable> allDT)
    {
        Dictionary<DateMY, DataTable> dictDT = new();
        foreach (DataTable dt in allDT)
        {
            string my_str = dt.Rows[0]["ns1:DocDate"].ToString(); // Колонка с датой документа
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
        return dictDT;
    }

    /// <summary>
    /// Конвертация и сохранение каждого элемента словаря
    /// </summary>
    /// <param name="dictDT">Словарь DateTable</param>
    private static void ConvAndSave(Dictionary<DateMY, DataTable> dictDT)
    {
        int counter = 1; // Счетчик для прогресса
        string tt = Console.Title;

        foreach (KeyValuePair<DateMY, DataTable> dt in dictDT)
        {
            Console.Title = tt + $"DataTable: {dt.Key.GetDate()} | Прогресс = {Math.Round((double)(counter * 100 / dictDT.Count))}% [{counter}/{dictDT.Count}]";

            if ((int)INFO > 0) Console.WriteLine($"\nУдаление дублей. Ключ таблицы: {dt.Key.GetDate()}");
            dt.Value.AsEnumerable().Distinct(DataRowComparer.Default); // Удаление дублей

            using ExcelUse excel = new(INFO);
            excel.Convert(dt.Value);
            excel.SaveAs(dirOut, GetFileName(directory, dt.Key.month, dt.Key.year));
            counter++;
        }
    }

    #endregion

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