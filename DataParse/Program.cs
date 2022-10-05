using System;
using System.Collections.Generic;
using System.Linq;
using System.IO;
using System.Diagnostics;
using System.Text.RegularExpressions;
using OfficeOpenXml;


namespace DataParse
{
    class Program
    {
        static void Main(string[] args)
        {
            bool isRepeat;
            do
            {
                isRepeat = false;
                Console.Write("Введите путь к файлу: ");
                string fileName = Console.ReadLine();
                Console.WriteLine("Создание отчета...");
                List<string> lines = File.ReadAllLines(fileName).ToList();

                ExcelPackage package = new ExcelPackage();
                ExcelPackage.LicenseContext = LicenseContext.Commercial;
                ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Отчет");

                sheet.Cells["A1"].Value = "Логин";
                sheet.Cells["B1"].Value = "Компьютер";
                sheet.Cells["C1"].Value = "Программа";
                sheet.Cells["D1"].Value = "Заголовок";
                sheet.Cells["E1"].Value = "Дата и время";

                int index = 2;
                lines.ForEach(i =>
                {
                    string date = i.GetString(0, i.IndexOf("__"));
                    string time = i.GetString(i.IndexOf("__") + 2, i.IndexOf("^^^")).Replace('-', ':');
                    time = time.GetString(0, time.LastIndexOf(':'));

                    DateTime myDate = DateTime.ParseExact($"{date} {time}", "yyyy-MM-dd HH:mm:ss",
                                           System.Globalization.CultureInfo.InvariantCulture);

                    string userLogin = new Regex(@"\^{3}\w*\^{3}").Match(i).Value.Replace("^", "");
                    string computer = new Regex(@"\^{3}\w*\^{6}").Match(i).Value.Replace("^", "");
                    string programName = new Regex(@"\^{6}.*\^{3}").Match(i).Value.Replace("^", "");
                    string programHeader = new Regex(@"\^{3}.+?#{4}").Match(i).Value;
                    programHeader = programHeader.GetString(programHeader.LastIndexOf("^^^"), programHeader.Length).Replace("^", "").Replace("#", "");

                    sheet.Cells[$"A{index}"].Value = userLogin;
                    sheet.Cells[$"B{index}"].Value = computer;
                    sheet.Cells[$"C{index}"].Value = programName;
                    sheet.Cells[$"D{index}"].Value = programHeader;
                    sheet.Cells[$"E{index}"].Value = $"{date} {time}";
                    index++;
                });

                try
                {
                    File.WriteAllBytes("Отчет по сотрудникам.xlsx", package.GetAsByteArray());
                    Console.WriteLine("Отчет создан.Открыть?y/n");
                    ConsoleKey key = Console.ReadKey(true).Key;
                    if (key == ConsoleKey.Y)
                    {
                        Process.Start("Отчет по сотрудникам.xlsx");
                    }
                }
                catch (IOException e)
                {
                    Console.WriteLine("Ошибка при сохранении отчета.Возможно отчет уже открыт");
                }
                finally
                {
                    Console.WriteLine("Перезапустить программу?y/n");
                    ConsoleKey key = Console.ReadKey(true).Key;
                    if (key == ConsoleKey.Y)
                    {
                        isRepeat = true;
                    }
                }

                Console.Clear();
            }
            while (isRepeat);
        }
    }

    static class StringExtension
    {
        public static string GetString(this string self, int oi, int ni)
        {
            string result = "";
            for (int i = oi; i < ni; i++)
            {
                result += self[i];
            }
            return result;
        }
    }
}
