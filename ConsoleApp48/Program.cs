using OfficeOpenXml;
using System.Formats.Asn1;

using CsvHelper;
using System.Globalization;
using Microsoft.VisualBasic.FileIO;
using System.Reflection.PortableExecutable;
using System.Text;
using System.Diagnostics;
using System.Threading;
using System.Collections.Concurrent;
using static System.Net.Mime.MediaTypeNames;
using System.Collections;

namespace ConsoleApp48
{
    internal class Program
    {
        static void Green()
        {
            Console.ForegroundColor = ConsoleColor.Green;
        }

        static void Gray()
        {
            Console.ForegroundColor = ConsoleColor.Gray;
        }
        static void Red()
        {
            Console.ForegroundColor = ConsoleColor.Red;
        }
        //Самое производительное
        static void Main(string[] arg)
        {
            try
            {
                if (arg.Length == 0)
                {
          
                    Exception exc = new Exception($"У программы 2 параметра при запуске!");
                    exc.HelpLink = "Параметр 1 :путь файла .csv \nПараметр 2 : куда сохранить и указать файл с путем.xslx";
                    exc.Data.Add("Время возникновения", DateTime.Now);
                    exc.Data.Add("Причина", $"Параметр не передан путь файла .csv и путь сохранения и название файла.xlsx");

                
                  
                    throw exc;

                }
                Console.WriteLine($"Параметр 1 :{arg[0]} ");
                Console.WriteLine($"Параметр 2 : {arg[1]} ");
                Stopwatch stopwatch = new Stopwatch();
                stopwatch.Start();
                // Установите путь к выходной директории для XLSX файлов
                string xlsxDirectoryPath = Environment.CurrentDirectory.ToString() + "\\Test";
                string[] csvFiles = new string[1];
                csvFiles[0] = arg[0];
                // Создание задачи для каждого CSV файла
                Task[] tasks = new Task[csvFiles.Length];
                for (int i = 0; i < csvFiles.Length; i++)
                {
                    string csvFilePath = csvFiles[i];
                    string xlsxFilePath = arg[1];
                    tasks[i] = Task.Run(() => ConvertCsvToXlsx(csvFilePath, xlsxFilePath));
                }
                // Ожидание завершения всех задач
                Task.WaitAll(tasks);
                Console.WriteLine("Преобразование завершено успешно!");
                stopwatch.Stop();
                TimeSpan timeSpan = stopwatch.Elapsed;
                Console.WriteLine("Время" + timeSpan);
                Environment.Exit(0);

            }
            catch (Exception ex)
            {
       
                Green();
                Console.WriteLine(ex.Message.ToString());
                Console.WriteLine(ex.HelpLink + "\n\n");
                Gray();
                Environment.Exit(0);

            }
        }

        public static void ConvertCsvToXlsx(string csvFilePath, string xlsxFilePath)
        {
            try { 
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

                using (ExcelPackage package = new ExcelPackage())
                {
                    ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Sheet1");

                    // Чтение CSV файла и запись данных в Excel
                    using (StreamReader reader = new StreamReader(csvFilePath, Encoding.GetEncoding("windows-1251")))
                    {
                        string firstLine = reader.ReadLine();

                        string[] fields = firstLine.Split(';', ',');

                        for (int col = 0; col < fields.Length; col++)
                        {
                            worksheet.Cells[1, col + 1].Value = fields[col]?.Trim('"', '.', ',');
                        }

                        int currentRow = 2;

                        while (!reader.EndOfStream)
                        {
                            string line = reader.ReadLine();
                            fields = line.Split(';', ',');
                            for (int col = 0; col < fields.Length; col++)
                            {
                                worksheet.Cells[currentRow, col + 1].Value = fields[col]?.Trim('"', '.', ',');
                            }

                            currentRow++;
                        }
                    }

                    FileInfo xlsxFile = new FileInfo(xlsxFilePath);
                    package.SaveAs(xlsxFile);
                }
            }
            catch(Exception ex)
            {
                Console.WriteLine(ex.Message.ToString());
                Console.WriteLine(ex.HelpLink + "\n\n");
                Environment.Exit(0);

            }
        }
    }
}

