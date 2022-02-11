using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace CompareExcel;

/// <summary>
/// Кол-во данных для вывода в консоль
/// </summary>
public enum INFOConsole
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

/// <summary>
/// Класс для и вывода информации в консоль
/// </summary>
public static class CConsole
{
    /// <summary>
    /// Уровень вывода информации в консоль (по умолчанию Main)
    /// </summary>
    internal static INFOConsole INFO = INFOConsole.Main;

    private static readonly char[] spinArr = { '|', '/', '-', '\\' };
    /// <summary>
    /// Индекс массива спинера
    /// </summary>
    private static int spinIndex = 0;

    // Для всех методов уровень информирования по умолчанию Main, т.к. большинство выводов именно такого уровня

    /// <summary>
    /// Пишет текст указанным цветом
    /// </summary>
    /// <param name="str">Текст для вывода в консоль</param>
    /// <param name="col">Цвет которым будет написан текст</param>
    /// <param name="newLine">Добавлять новую строку (WriteLine)?</param>
    /// <param name="infLVL">Уровень информирования этой строчки. 0 - None, 1 - Main, 2 - All</param>
    public static void Print(string str, INFOConsole strInfo = INFOConsole.Main, ConsoleColor col = ConsoleColor.White, bool newLine = true)
    {
        // Общий ур-вень информирования должен быть выше или равен strInfo
        if (INFO < strInfo) return;

        ConsoleColor def = Console.ForegroundColor; // Получение стандартного цвета

        Console.ForegroundColor = col;

        if (newLine) Console.WriteLine(str);
        else Console.Write(str);

        Console.ForegroundColor = def; // Возвращение стандартного цвета
    }
    /// <summary>
    /// Разделитель информации в строке
    /// </summary>
    /// <param name="lenght">Длина деления</param>
    public static void Separator(int lenght, INFOConsole strInfo = INFOConsole.Main)
    {
        if (INFO < strInfo) return;

        Console.WriteLine();
        for (int i = 0; i < lenght; i++)
            Console.Write("#");
        Console.WriteLine();
    }

    /// <summary>
    /// Разделитель информации в строке
    /// </summary>
    /// <param name="min">Минимальное кол-во символов</param>
    /// <param name="max">Максимальное кол-во символов</param>
    public static void Separator(int min = 30, int max = 50, INFOConsole strInfo = INFOConsole.Main)
    {
        if (INFO < strInfo) return;

        Random rnd = new();
        int ss = rnd.Next(min, max);
        Separator(ss, strInfo);
    }

    /// <summary>
    /// Пишет в консоль с той же строчки (перезаписывает)
    /// </summary>
    /// <param name="str">Строка</param>
    /// <param name="horiaontal">Горизонтальное смещение</param>
    public static void SetLine(string str, INFOConsole strInfo = INFOConsole.Main, int horiaontal = 0)
    {
        if (INFO != strInfo) return;

        Console.SetCursorPosition(horiaontal, Console.GetCursorPosition().Top);
        Console.Write(str);
    }

    /// <summary>
    /// Отображает прогресс в виде: {Строка 1}8% {Строка 2}[80/1000]
    /// </summary>
    /// <param name="index">Индекс текущего элемента в цикле</param>
    /// <param name="maxCount">Максимальное кол-во элементов</param>
    /// <param name="str">Строка 1</param>
    /// <param name="str2">Строка 2</param>
    /// <param name="spiner">Отображать спинер?</param>
    /// <returns></returns>
    public static string Progress(int index, int maxCount, string str = "", string str2 = "", bool spiner = false)
    {
        return $"{str}{Math.Round((double)((index + 1) * 100 / maxCount))}% " +
            $"{str2}[{index + 1} / {maxCount}]" +
            $"{((spiner && index < maxCount - 1) ? " " + GetSpiner() : "  ")}";
    }

    /// <summary>
    /// Возвращает строку прогресса в виде визуальной полосы
    /// </summary>
    /// <param name="index">Текущий индекс цикла</param>
    /// <param name="maxCount">Максимальное кол-во</param>
    /// <param name="str">Строка до</param>
    /// <param name="lenght">Длина визуальной полосы</param>
    /// <returns></returns>
    public static string ProgressBar(int index, int maxCount, string str = "", int lenght = 20, bool spiner = false)
    {
        string sout = $"{str}|";
        double percent = Math.Round((double)((index + 1) * 100 / maxCount));

        for (int i = 0; i < Math.Round(lenght * (percent) / 100, MidpointRounding.ToZero); i++) sout += "█";
        if(percent != 100) sout += (percent % 10 >= 5) ? "▌" : " "; // Если процентов >=x5
        for (int i = sout.Length - 1; i < str.Length + lenght; i++) sout += " ";

        sout += $"| {percent}% ";
        sout += $"[{index + 1} / {maxCount}]";
        sout += $"{((spiner && index < maxCount - 1) ? " " + GetSpiner() : "  ")}";
        return sout;
    }

    /// <summary>
    /// Дает анимацию спинера
    /// </summary>
    public static char GetSpiner()
    {
        spinIndex = (spinIndex < spinArr.Length - 1) ? spinIndex + 1 : 0;
        return spinArr[spinIndex];
    }

    /// <summary>
    /// Новая строчка
    /// </summary>
    /// <param name="strInfo"></param>
    public static void WriteLine(INFOConsole strInfo = INFOConsole.Main)
    {
        if (INFO != strInfo) return;

        Console.WriteLine();
    }

    /// <summary>
    /// Ждет пока пользователь нажмет клавишу
    /// </summary>
    /// <param name="strInfo"></param>
    public static void Await(INFOConsole strInfo = INFOConsole.Main)
    {
        if (INFO != strInfo) return;

        Console.WriteLine("Нажмите любую клавишу...");
        Console.ReadKey();
    }
}

