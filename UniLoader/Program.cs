using System;
using System.Data;
using System.IO;
using System.Linq;
using UniLoader.Config;
using UniLoader.DataClients;
using UniLoader.DataParsers;

namespace UniLoader
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.SetWindowSize(140, 30);
            string message;
            int consoleCursor = 0;
            var configPath = Directory.GetCurrentDirectory() + "\\appconfig.json";
            //Ставим логгер
            if (!Directory.Exists("\\logs")) Directory.CreateDirectory("logs");
            Logger.SetFile(Directory.GetCurrentDirectory() + "\\logs\\" + $"user_log_{DateTime.Now:ddMMyy_hhmm}.txt");

            //Чтение конфига
            Console.WriteLine("Подготовка загрузки.\nЧтение конфигурационного файла...");
            try
            {
                ConfigUtil configUtil = new JsonConfigUtil(configPath);
                configUtil.ReadConfig();

                message = $"Конфигурационный файл {configPath} успешно загружен.";
                WriteLineColor(message, ConsoleColor.DarkGreen);
                Logger.WriteLine(message);
            }
            catch (Exception)
            {
                message = $"Ошибка загрузки конфигурационного файла {configPath}!";
                WriteLineColor(message, ConsoleColor.DarkRed);
                Logger.WriteLine(message);
                SayByeBye();
                return;
            }

            DataClient dataClient = new OracleDataClient(AppConfig.DatabaseConnectionString); //<--для пластичности можно переделать в свитч

            if (!Directory.Exists(AppConfig.WorkingDirectory) || !dataClient.TestConnection())
            {
                message = "Ошибка: Загрузка невозможна (отсутствует рабочая директория и/или соединение с БД).";
                WriteLineColor(message, ConsoleColor.DarkRed);
                Logger.WriteLine(message);
                SayByeBye();
                return;
            }
            foreach (var table in AppConfig.Tables)
            {
                dataClient.ConfTable = table;
                if (Directory.Exists(AppConfig.WorkingDirectory + table.SubDirectory))
                {
                    consoleCursor += 6;
                    var timer = new System.Diagnostics.Stopwatch();
                    timer.Start();

                    DrawLine();
                    WriteLineDifferentColor("Обработка поддиректории {1}", $"{table.SubDirectory}", ConsoleColor.DarkYellow);
                    Logger.WriteLine($"Обработка поддиректории {table.SubDirectory}");

                    var pathList = Directory.GetFiles(AppConfig.WorkingDirectory + table.SubDirectory + "\\",
                        AppConfig.FileExtension).ToList();

                    WriteLineDifferentColor("Будет обработано {1} файлов.\nЗагрузка файлов:\n", $"{pathList.Count}", ConsoleColor.DarkYellow);
                    Logger.WriteLine($"Будет обработано {pathList.Count} файлов.\nЗагрузка файлов:\n");

                    foreach (var path in pathList)
                    {
                        ClearLine(consoleCursor);
                        ClearLine(consoleCursor + 1);
                        Console.SetCursorPosition(0, consoleCursor);

                        WriteLineDifferentColor("Прогресс обработки ({1}/{2}):", $"{pathList.IndexOf(path) + 1}",
                            $"{pathList.Count}", ConsoleColor.DarkGreen, ConsoleColor.DarkYellow);

                        message = $"Файл {path.Split('\\')[path.Split('\\').Length - 1]} ";
                        Console.Write(message);
                        Logger.WriteLine(message);

                        try
                        {
                            DataParser parser = new XlsxDataParser(path, table);
                            DataTable dataTable = parser.ReadFile();
                            dataClient.Send(dataTable);

                            if (AppConfig.DeleteFilesAfterLoad) File.Delete(path);
                        }
                        catch (Exception)
                        {
                            Console.WriteLine("ОШИБКА.");
                            Logger.WriteLine("ОШИБКА. Файл не загружен.");
                        }
                    }
                    timer.Stop();
                    message = $"Файлы поддиректории загружены за {timer.Elapsed.TotalMinutes} минут.";
                    WriteLineColor(message, ConsoleColor.DarkGreen);
                    Logger.WriteLine(message);
                }
            }

            WriteLineColor("Все файлы успешно загружены.", ConsoleColor.Green);
            Logger.WriteLine("Все файлы успешно загружены.");
            SayByeBye();
        }

        public static void SayByeBye()
        {
            Console.WriteLine("Программа завершает работу. Нажмите любую клавишу.");
            Console.ReadLine();
        }

        public static void DrawLine()
        {
            Console.ForegroundColor = ConsoleColor.Blue;
            Console.WriteLine("//================================================//");
            Console.ForegroundColor = ConsoleColor.Gray;
        }

        public static void WriteLineDifferentColor(string message, string colored, ConsoleColor color)
        {
            var strings = message.Split(new[] { "{1}" }, StringSplitOptions.None);

            Console.Write(strings[0]);
            Console.ForegroundColor = color;
            Console.Write(colored);
            Console.ResetColor();
            Console.WriteLine(strings[1]);
        }

        public static void WriteLineDifferentColor(string message, string colored1, string colored2,
            ConsoleColor color1, ConsoleColor color2)
        {
            var strings = message.Split(new[] { "{1}", "{2}" }, StringSplitOptions.None);

            Console.Write(strings[0]);
            Console.ForegroundColor = color1;
            Console.Write(colored1);
            Console.ResetColor();
            Console.Write(strings[1]);
            Console.ForegroundColor = color2;
            Console.Write(colored2);
            Console.ResetColor();
            Console.WriteLine(strings[2]);
        }

        public static void WriteLineColor(string message, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(message);
            Console.ResetColor();
        }

        public static void WriteColor(string message, ConsoleColor color)
        {
            Console.ForegroundColor = color;
            Console.Write(message);
            Console.ResetColor();
        }

        public static void ClearLine(int position)
        {
            Console.SetCursorPosition(0, position);
            for (int i = 0; i < Console.WindowWidth; i++) Console.Write(" ");
        }
    }
}
