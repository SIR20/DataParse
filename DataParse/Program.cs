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
           
            Console.Write("Введите путь к файлу: ");
            string fileName = Console.ReadLine();
            CreateReport(fileName);
        }

        static void CreateReport(string fileName)
        {
            string reportFileName = Path.GetDirectoryName(fileName) + @"\" + Path.GetFileNameWithoutExtension(fileName) + ".xlsx";
            Dictionary<string, App> apps = new Dictionary<string, App>();

            ExcelPackage package = new ExcelPackage();
            ExcelPackage.LicenseContext = LicenseContext.Commercial;
            ExcelWorksheet sheet = package.Workbook.Worksheets.Add("Отчет");

            Console.WriteLine("Создание отчета...");

            sheet.Cells["A1"].Value = "Программа";
            sheet.Columns[1].Width = 35;
            

            sheet.Cells["B1"].Value = "Время";
            sheet.Columns[2].Width = 18;

            sheet.Cells["C1"].Value = "Итого:";
            sheet.Columns[3].Width = 17;

            List<string> lines = File.ReadAllLines(fileName).ToList();
            lines.ForEach(i =>
            {
                string date = i.GetString(0, i.IndexOf("__"));
                string time = i.GetString(i.IndexOf("__") + 2, i.IndexOf("^^^")).Replace('-', ':');
                time = time.GetString(0, time.LastIndexOf(':'));
                DateTime openTime = DateTime.ParseExact($"{date} {time}", "yyyy-MM-dd HH:mm:ss",
                                       System.Globalization.CultureInfo.InvariantCulture);

                string programName = new Regex(@"\^{6}.*\^{3}").Match(i).Value.Replace("^", "");
                string programHeader = new Regex(@"\^{3}.+?#{4}").Match(i).Value;
                programHeader = programHeader.GetString(programHeader.LastIndexOf("^^^"), programHeader.Length).Replace("^", "").Replace("#", "");

                if (apps.ContainsKey(programName))
                {
                    apps[programName].AddOpenTime(programHeader, openTime);
                }
                else
                {
                    App app = new App(programName);
                    app.AddOpenTime(programHeader, openTime);
                    apps.Add(programName, app);
                }
            });

            int allTimeSecond = 0;
            int index = 3;
            apps.Keys.ToList().ForEach(programName =>
            {
                
                int timeSecond = 0;
                List<OpenTime> openTimes = apps[programName].OpenTimes;
                openTimes.ForEach(j =>
                {
                    timeSecond += j.Time.Second;
                });
                TimeSpan time = TimeSpan.FromSeconds(timeSecond);

                sheet.Cells[$"A{index}"].Value = programName;
                sheet.Cells[$"B{index}"].Value = $"{time.Days}дн.{time.Hours}ч.{time.Minutes}м.{time.Seconds}с.";

                allTimeSecond += timeSecond;
                index++;
            });

            TimeSpan allTime = TimeSpan.FromSeconds(allTimeSecond);
            sheet.Cells[$"C2"].Value = $"{allTime.Days}дн.{allTime.Hours}ч.{allTime.Minutes}м.{allTime.Seconds}с.";

            try
            {
                File.WriteAllBytes(reportFileName, package.GetAsByteArray());
                Console.WriteLine("Отчет создан.Открыть?y/n");
                ConsoleKey key = Console.ReadKey(true).Key;
                if (key == ConsoleKey.Y)
                {
                    Process.Start(reportFileName);
                }
            }
            catch (IOException)
            {
                Console.WriteLine("Ошибка при сохранении отчета.Возможно отчет уже открыт");
            }
        }
    }

    class App
    {
        public App(string name)
        {
            Name = name;
            OpenTimes = new List<OpenTime>();
        }

        public string Name { get; set; }
        public List<OpenTime> OpenTimes { get; set; }

        public void AddOpenTime(string header, DateTime time) => OpenTimes.Add(new OpenTime(header, time));
    }

    class OpenTime
    {
        public OpenTime(string header, DateTime time)
        {
            Header = header;
            Time = time;
        }

        public string Header { get; set; }
        public DateTime Time { get; set; }
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
