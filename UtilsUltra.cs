using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;
using System.Threading;

public static class Utils
{
    /*
     *              _____                _____                    _____                    _____            _____          
     *             /\    \              /\    \                  /\    \                  /\    \          /\    \         
     *            /::\____\            /::\    \                /::\    \                /::\____\        /::\    \        
     *           /:::/    /            \:::\    \               \:::\    \              /:::/    /       /::::\    \       
     *          /:::/    /              \:::\    \               \:::\    \            /:::/    /       /::::::\    \      
     *         /:::/    /                \:::\    \               \:::\    \          /:::/    /       /:::/\:::\    \     
     *        /:::/    /                  \:::\    \               \:::\    \        /:::/    /       /:::/__\:::\    \    
     *       /:::/    /                   /::::\    \              /::::\    \      /:::/    /        \:::\   \:::\    \   
     *      /:::/    /      _____        /::::::\    \    ____    /::::::\    \    /:::/    /       ___\:::\   \:::\    \  
     *     /:::/____/      /\    \      /:::/\:::\    \  /\   \  /:::/\:::\    \  /:::/    /       /\   \:::\   \:::\    \ 
     *    |:::|    /      /::\____\    /:::/  \:::\____\/::\   \/:::/  \:::\____\/:::/____/       /::\   \:::\   \:::\____\
     *    |:::|____\     /:::/    /   /:::/    \::/    /\:::\  /:::/    \::/    /\:::\    \       \:::\   \:::\   \::/    /
     *     \:::\    \   /:::/    /   /:::/    / \/____/  \:::\/:::/    / \/____/  \:::\    \       \:::\   \:::\   \/____/ 
     *      \:::\    \ /:::/    /   /:::/    /            \::::::/    /            \:::\    \       \:::\   \:::\    \     
     *       \:::\    /:::/    /   /:::/    /              \::::/____/              \:::\    \       \:::\   \:::\____\    
     *        \:::\__/:::/    /    \::/    /                \:::\    \               \:::\    \       \:::\  /:::/    /    
     *         \::::::::/    /      \/____/                  \:::\    \               \:::\    \       \:::\/:::/    /     
     *          \::::::/    /                                 \:::\    \               \:::\    \       \::::::/    /      
     *           \::::/    /                                   \:::\____\               \:::\____\       \::::/    /       
     *            \__/____/                                     \::/    /                \::/    /        \::/    /        
     *                                                           \/____/                  \/____/          \/____/         
     *                                                                                                                     
     */
    private static Stopwatch executionTime = new Stopwatch();
    public static Random random = new Random();
    public static ProgramInfo info;

    /// <summary>
    /// Вывести авторов + время выполнение программы
    /// </summary>
    public static void PrintAuthors()
    {
        Console.Title = info.name;
        //Authors
        Console.ForegroundColor = ConsoleColor.DarkGray;
        Console.WriteLine("----------------------");
        Console.WriteLine($"Задание #{info.name.Split('.')[0]}");
        Console.WriteLine($"Выполнил {info.author}");
        Console.WriteLine("ИСП211       ");
        Console.WriteLine("----------------------");
        Console.ResetColor();

        //OTHER
        executionTime.Start();
        //on exit
        AppDomain.CurrentDomain.ProcessExit += delegate (object sender, EventArgs e)
        {
            executionTime.Stop();
            Console.ForegroundColor = ConsoleColor.DarkGray;
            Console.WriteLine("\n----------------------");
            Console.WriteLine($"Время: {Math.Round(executionTime.Elapsed.TotalMilliseconds)}мс | {Math.Round(executionTime.Elapsed.TotalSeconds, 2)}с");
        };

        //HANDLE F1
        WriteLineColored("\n<fg=gray>Нажмите <fg=darkyellow>F1 <fg=gray>для вывода информации о программе\n<fg=darkyellow>Любой символ<fg=gray> для продолжения");
        Console.ForegroundColor = ConsoleColor.Black;
        if (Console.ReadKey().Key == ConsoleKey.F1)
        {
            string accent = "<r><fg=green>";
            string normal = "<r><fg=cyan>";
            WriteLineColored($"\n{accent}Автор:{normal} " + info.author);
            WriteLineColored($"{accent}Название:{normal} " + info.name);
            WriteLineColored($"{accent}Описание:{normal} " + info.description);
            WriteLineColored($"{accent}Инструкция:{normal} " + info.instruction + "\n");
        }
        else
        {
            Console.WriteLine();
        }
        Console.ResetColor();
    }

    /// <summary>
    /// Быстрый ввод переменной + парсинг в переданный тип
    /// </summary>
    /// <typeparam name="T">Тип переменной</typeparam>
    /// <param name="name">Имя переменной (опционально)</param>
    /// <returns>Ввод с консоли</returns>
    public static T Input<T>(string name = "?")
    {
        WriteColored($"Введите {name}<r>:\n");
        return (T)Convert.ChangeType(Console.ReadLine(), typeof(T));
    }
    public static void Input<T>(this T variable, string name = "?")
    {
        WriteColored($"Введите {name}<r>:\n");
        variable = (T)Convert.ChangeType(Console.ReadLine(), typeof(T));
    }

    /// <summary>
    /// Быстрый вывод переменной
    /// </summary>
    /// <typeparam name="T">Тип переменной (опционально, можно вызвать с переменной)</typeparam>
    /// <param name="variable">Переменная</param>
    /// <param name="name">Имя (опционально)</param>
    public static void Output<T>(this T variable, string name = "")
    {
        WriteLineColored($"{name.Replace(':', '\0')}<r> = {variable}");
    }
    public static void Output<T>(this T[] variable, string name = "")
    {
        WriteLineColored($"{name.Replace(':', '\0')}<r> = {String.Join(", ", variable)}");
    }
    public static void OutputArray(this Array array, bool printIndices = false)
    {
        int maxLength = 0;
        array.Map(element => { maxLength = Math.Max(maxLength, element.ToString().Length); });
        if (array.Rank == 2)
        {
            for (int y = 0; y < array.GetLength(0); y++)
            {
                for (int x = 0; x < array.GetLength(1); x++)
                {
                    if (printIndices)
                    {
                        WriteColored($"<fg=DarkGray>[{y},{x}]:<r>{array.GetValue(y, x)}");
                    }
                    else
                    {
                        Console.Write(array.GetValue(y, x));
                    }
                    for (int i = 0; i <= maxLength - array.GetValue(y, x).ToString().Length + 1; i++) Console.Write(" ");
                }
                Console.WriteLine();
            }
        }
        else
        {
            for (int x = 0; x < array.GetLength(0); x++)
            {
                if (printIndices)
                {
                    WriteColored($"<fg=DarkGray>[{x}]:<r>{array.GetValue(x)}");
                }
                else
                {
                    Console.Write(array.GetValue(x));
                }
                for (int i = 0; i <= maxLength - array.GetValue(x).ToString().Length + 1; i++) Console.Write(" ");
            }
        }
        Console.WriteLine();
    }

    /// <summary>
    /// Вывод красивой ошибки
    /// </summary>
    /// <param name="errorMessage">Содержание ошибки</param>
    public static void PrintError(string errorMessage)
    {
        //Console.BackgroundColor = ConsoleColor.DarkRed;
        //Console.WriteLine($"ошибка: {errorMessage}".ToUpper());
        //Console.ResetColor();
        WriteLineColored($"<bg=DarkRed>ошибка: {errorMessage}".ToUpper());
    }


    /// <summary>
    /// Выполнение действие над массивом любой размерности
    /// </summary>
    /// <param name="array">Массив (можно вызвать из него)</param>
    /// <param name="action">Само действие</param>
    /// <param name="dimension"></param>
    /// <param name="indices"></param>
    public static void Map(this Array array, Func<object, object> action, int dimension = 0, int[] indices = null)
    {
        if (indices == null) indices = new int[array.Rank];

        for (int i = 0; i < array.GetLength(dimension); i++)
        {
            indices[dimension] = i;

            if (dimension == array.Rank - 1)
            {
                var newValue = action(array.GetValue(indices));
                array.SetValue(newValue, indices);
            }
            else
            {
                Map(array, action, dimension + 1, indices);
            }
        }
    }
    public static void Map(this Array array, Action<object> action, int dimension = 0, int[] indices = null)
    {
        if (indices == null) indices = new int[array.Rank];

        for (int i = 0; i < array.GetLength(dimension); i++)
        {
            indices[dimension] = i;

            if (dimension == array.Rank - 1)
            {
                action(array.GetValue(indices));
            }
            else
            {
                Map(array, action, dimension + 1, indices);
            }
        }
    }

    /// <summary>
    /// Возвращает случайный элемент из массива или строки
    /// </summary>
    /// <typeparam name="T">Тип переменной (можно вызвать из массива)</typeparam>
    /// <param name="array">Массив (можно вызвать из массива)</param>
    /// <returns>Случайный элемент из массива</returns>
    public static T PickRandom<T>(this Array array)
    {
        List<T> elements = new List<T>();
        array.Map(item => elements.Add((T)item));
        return elements[random.Next(elements.Count)];
    }
    public static string PickRandom(this string str)
    {
        return str[random.Next(str.Length)].ToString();
    }

    /// <summary>
    /// Массив [][] => Массив [,]
    /// </summary>
    /// <typeparam name="T">Тип массива (можно вызвать из массива)</typeparam>
    /// <param name="array">массив</param>
    /// <returns>Массив [,]</returns>
    public static T[,] ToArray<T>(this T[][] array)
    {
        int rows = array.Length;
        int cols = array[0].Length;
        T[,] rectArray = new T[rows, cols];

        for (int y = 0; y < rows; y++)
        {
            for (int x = 0; x < cols; x++)
            {
                rectArray[y, x] = array[y][x];
            }
        }

        return rectArray;
    }

    /// <summary>
    /// Вывод цветного текста. <fg=color> - цвет текста, <bg=color> - цвет фона, <r> - сброс цвета
    /// </summary>
    /// <param name="text">Текст</param>
    public static void WriteColored(string text, bool resetColor = true)
    {
        if (resetColor) text += "<r>";
        string pattern = @"(<[^>]+>)|([^<]+)";
        List<string> wordList = new List<string>();

        foreach (Match match in Regex.Matches(text, pattern))
        {
            if (match.Groups[1].Success)
            {
                wordList.Add(match.Groups[1].Value);
            }
            else if (match.Groups[2].Success)
            {
                wordList.Add(match.Groups[2].Value);
            }
        }
        Array words = wordList.ToArray();
        words.Map(item =>
        {
            var itemString = item.ToString();
            if (itemString[0] == '<' && itemString[itemString.Length - 1] == '>')
            {
                //tag
                itemString = string.Concat(itemString.Where(c => !char.IsWhiteSpace(c))).ToLower();
                if (itemString.Contains("="))
                {
                    var parts = itemString.Split('=');
                    ConsoleColor color = new ConsoleColor();
                    foreach (ConsoleColor c in (ConsoleColor[])Enum.GetValues(typeof(ConsoleColor)))
                    {
                        if (c.ToString().ToLower() + ">" == parts[1])
                        {
                            color = c;
                        }
                    }
                    if (parts[0] == "<bg")
                    {
                        Console.BackgroundColor = color;
                    }
                    else if (parts[0] == "<fg")
                    {
                        Console.ForegroundColor = color;
                    }
                    else
                    {
                        Console.Write(item);
                    }
                }
                else
                {
                    if (itemString == "<r>")
                    {
                        Console.ResetColor();
                    }
                    else
                    {
                        Console.Write(item);
                    }
                }
            }
            else
            {
                //word
                Console.Write(item);
            }
        });
    }
    public static void WriteLineColored(string text)
    {
        WriteColored(text + "\n");
    }

    /// <summary>
    /// struct с полезными константами
    /// </summary>
    public struct Constants
    {
        /// <summary>
        /// abcdefghijklmnopqrstuvwxyz
        /// </summary>
        public static string lowerAlphabet => "abcdefghijklmnopqrstuvwxyz";
        /// <summary>
        /// ABCDEFGHIJKLMNOPQRSTUVWXYZ
        /// </summary>
        public static string upperAlphabet => lowerAlphabet.ToUpper();
        /// <summary>
        /// 0123456789
        /// </summary>
        public static string numbers => "0123456789";
        /// <summary>
        /// abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ
        /// </summary>
        public static string fullAlphabet => lowerAlphabet + upperAlphabet;
        /// <summary>
        /// 0123456789abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ
        /// </summary>
        public static string alpabetNumbers => numbers + fullAlphabet;
    }

    /// <summary>
    /// Информация о текущей программе
    /// </summary>
    public class ProgramInfo
    {
        public string author;
        public string name;
        public string description;
        public string instruction;

        public ProgramInfo(string author, string name, string description, string instruction)
        {
            this.author = author;
            this.name = name;
            this.description = description;
            this.instruction = instruction;
        }
    }
}