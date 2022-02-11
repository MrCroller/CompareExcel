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
    /// Указывает начать ли с последнего скомпелированного файла
    /// </summary>
    static private bool last_Dir = true;
    /// <summary>
    /// Директория вывода файлов
    /// </summary>
    static private string dirOut;
    /// <summary>
    /// Название консоли
    /// </summary>
    static private string title = "Excel Merge ";
    /// <summary>
    /// Кол-во повторных попыток вызова ошибки
    /// </summary>
    static private int errorReplayCount = 3;

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
            CConsole.Print(string.Join("\n", dirArr), INFOConsole.All);
            if (dirArr.Length == 0)
            {
                throw new Exception("Указанная папка не имеет вложенных папок");
            }

            int lastindex = 0; // индекс последнего скомпелированного файла
            if (new DirectoryInfo(dirOut).GetFiles().Length > 0 && last_Dir)
            {
                string lastFile = GetLastFile(dirOut);
                lastindex = Array.IndexOf(dirArr, $@"{directory}\{lastFile.Remove(lastFile.Length - 13)}"); // Название файла без даты в конце

                CConsole.Print($"fullFile: {lastFile}\n revert: {directory}\\{lastFile.Remove(lastFile.Length - 13)}\nlastindex: {lastindex}", INFOConsole.All);
                CConsole.Await(INFOConsole.All);
            }

            for (int i = lastindex; i < dirArr.Length; i++)
            {
                Console.Title = $"dir:[{i + 1}/{dirArr.Length}] ";

                List<DataTable> rr = ReadRange(dirArr[i]);
                Dictionary<DateMY, DataTable> dictDT = SortDT(rr);
                ConvAndSave(dictDT, dirArr[i]);
            }
        }
        else
        {
            List<DataTable> rr = ReadRange(directory);
            Dictionary<DateMY, DataTable> dictDT = SortDT(rr);
            ConvAndSave(dictDT, directory);
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
                    Console.WriteLine("-----Commands-----\n-help (показать команды)\n-d (Папка для чтения)\n-n (Учитывать вложенные папки): true - учитывать, false - не учитывать [по умолчанию]\n-o (Папка для вывода объединенного файла)\n-c (Ищет в директории выхода последний скомпелированный файл и начинает с него): true - начинать с последнего [по умолчанию], false - начинать с первой папки\n-i (Уровень информирования, выбор из 3 доступных: None - Без консольного вывода, Main - главные процессы [по умолчанию], All - Полный вывод, аж на каждую ячеечку. Много букав!)\nДля оболочки PowerShell оборачивайте путь в ковычки 'directory'\n------------------");
                    break;
                case "-d": // Папка для чтения
                    if (Directory.Exists(args[i + 1]))
                    {
                        directory = args[i + 1];
                    }
                    else
                    {
                        throw new Exception("Не задана папка для чтения или ее не существует");
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
                        CConsole.Print("Нераспознан параметр вложенных папок -d. Вывод по умолчанию: false", col: ConsoleColor.Red);
                    }
                    break;
                case "-o": // Директория выхода
                    if (Directory.Exists(args[i + 1]))
                    {
                        dirOut = args[i + 1];
                    }
                    else
                    {
                        throw new Exception("Не задана папка для вывода или ее не существует");
                    }
                    break;
                case "-c": // Начинать ли с последнего скомпелированного файла?
                    if (args[i + 1].ToLower() == "true")
                    {
                        last_Dir = true;
                    }
                    else if (args[i + 1].ToLower() == "false")
                    {
                        last_Dir = false;
                    }
                    else
                    {
                        CConsole.Print("Нераспознан параметр последнего файла. Вывод по умолчанию: true", col: ConsoleColor.Red);
                    }
                    break;
                case "-i": // Уровень информирования
                    switch (args[i + 1].ToLower())
                    {
                        case "none": CConsole.INFO = INFOConsole.None; break;
                        case "main": CConsole.INFO = INFOConsole.Main; break;
                        case "all": CConsole.INFO = INFOConsole.All; break;
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
        // Сканирует файлы формата Excel с {путь}\piev_{набор цифр}, для избежания временных файлов
        string[] filesDir = Directory.GetFiles(dirName, "piev_*.xlsx", SearchOption.TopDirectoryOnly);
        string folder = dirName.Split(@"\")[^1];
        string tt = Console.Title;

        List<DataTable> list = new();

        CConsole.Print($"\nСканирование директории: {dirName}");

        using ExcelUse excelApp = new();
        for (int i = 0; i < filesDir.Length; i++)
        {
            Console.Title = tt + $"Папка: {folder} | {CConsole.Progress(i, filesDir.Length, "Прогресс = ")}";
            CConsole.Print($"\n[{i + 1}/{filesDir.Length}]");

            DataTable dt = new();
            int error_Count = 0;
            do
            {
                CConsole.Print(filesDir[i], INFOConsole.All, ConsoleColor.Cyan);
                dt = excelApp.ReadFile(filesDir[i]);
                if (dt == null)
                {
                    error_Count++;
                    if (error_Count > errorReplayCount) continue;
                    CConsole.Print($"Повторная попытка прочтения файла {error_Count}");
                    Thread.Sleep(3000);
                }
            } while (dt == null && error_Count <= errorReplayCount);

            list.Add(dt);
        }
        Console.Title = title;

        return list;
    }

    /// <summary>
    /// Сортирует словарь таблиц по месяцам
    /// </summary>
    /// <param name="allDT">Коллекция для сортировки</param>
    /// <returns></returns>
    private static Dictionary<DateMY, DataTable> SortDT(List<DataTable> allDT)
    {
        CConsole.WriteLine();
        Dictionary<DateMY, DataTable> dictDT = new();
        string strsort = "Сортировка DataTable по датам: ";

        CConsole.Print(strsort, col: ConsoleColor.Red, newLine: false);
        for (int i = 0; i < allDT.Count; i++)
        {
            CConsole.SetLine(CConsole.Progress(i, allDT.Count, spiner: true), horiaontal: strsort.Length);

            string my_str = allDT[i].Rows[0]["ns1:DocDate"].ToString(); // Колонка с датой документа
            DateMY date = new()
            {
                month = Convert.ToInt32(my_str.ToString().Substring(3, 2)),
                year = Convert.ToInt32(my_str.ToString().Substring(6, 4))
            };

            if (dictDT.ContainsKey(date))
            {
                // Соединение таблиц. Сохранение данных и добавление других колонок
                dictDT[date].Merge(allDT[i], true, MissingSchemaAction.Add);
            }
            else
            {
                dictDT.Add(date, allDT[i]);
            }
        }
        CConsole.WriteLine();
        return dictDT;
    }

    /// <summary>
    /// Конвертация и сохранение каждого элемента словаря
    /// </summary>
    /// <param name="dictDT">Словарь DateTable</param>
    /// <param name="dir">Директория файла</param>
    private static void ConvAndSave(Dictionary<DateMY, DataTable> dictDT, string dir)
    {
        int counter = 1; // Счетчик для прогресса
        string tt = Console.Title;

        foreach (KeyValuePair<DateMY, DataTable> dt in dictDT)
        {
            Console.Title = tt + $"DataTable: {dt.Key.GetDate()} | {CConsole.Progress(counter, dictDT.Count, "Прогресс = ")}";
            CConsole.Print($"\nУдаление дублей. Ключ таблицы: {dt.Key.GetDate()}");
            dt.Value.AsEnumerable().Distinct(DataRowComparer.Default); // Удаление дублей

            using ExcelUse excel = new();
            excel.Convert(dt.Value);
            excel.SaveAs(dirOut, GetFileName(dir, dt.Key.month, dt.Key.year));
            counter++;
        }
    }

    #endregion

    #region Получение имен файлов

    /// <summary>
    /// Возвращает имя последнего созданного файла в директории
    /// </summary>
    /// <param name="DirFolder">Директория</param>
    /// <returns></returns>
    private static string GetLastFile(string DirFolder)
    {
        DateTime dt = new DateTime(1990, 1, 1);
        string fileName = "";

        FileSystemInfo[] fileSystemInfo = new DirectoryInfo(DirFolder).GetFileSystemInfos();
        foreach (FileSystemInfo fileSI in fileSystemInfo)
        {
            if (dt < Convert.ToDateTime(fileSI.CreationTime))
            {
                dt = Convert.ToDateTime(fileSI.CreationTime);
                fileName = fileSI.Name;
            }
        }
        return fileName;
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
        string mounthSTR = mounth < 10 ? "0" + mounth.ToString() : mounth.ToString();
        // Если путь с названием файла, то берется вторая подстрочка с конца (имя папки)
        return $@"{DirFolder.Split(@"\")[^(DirFolder.EndsWith(".xlsx") ? 2 : 1)]}_{mounthSTR}_{year}";
    }
    #endregion

    /// <summary>
    /// Остановка программы при ошибке
    /// </summary>
    /// <param name="er">Текст сообщения ошибки</param>
    private static void ErrorExit(string er)
    {
        CConsole.Print(er, INFOConsole.None, ConsoleColor.Red);
        CConsole.Await();
        Process.GetCurrentProcess().Kill();
    }
}