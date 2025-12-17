using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Reflection;
using ClosedXML.Excel;

namespace SystemsHighlighter.Tools
{
    public class PhysData
    {
        /// <summary>
        /// Код подсистемы (из колонки "Подсистема")
        /// </summary>
        public string PipeLine { get; set; }

        /// <summary>
        /// Тоннаж (из колонки "Тоннаж")
        /// </summary>
        public double Tonnage { get; set; }

        /// <summary>
        /// Протяженность (из колонки "Протяженность")
        /// </summary>
        public double Length { get; set; }

        /// <summary>
        /// Объем (из колонки "Объем")
        /// </summary>
        public double Volume { get; set; }

        /// <summary>
        /// Считывает все данные из указанного Excel-файла
        /// и возвращает список записей.
        /// </summary>
        /// <param name="fileName">Имя Excel-файла (по умолчанию "PhysData.xlsx")</param>
        /// <param name="sheetName">Имя листа в файле (по умолчанию "Data")</param>
        public static List<PhysData> LoadFromExcel(string filePath,
            string fileName = "PhysData.xlsx",
            string sheetName = "Data")
        {
            var result = new List<PhysData>();

            // Папка со сборкой
            //string assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            string fullPath = Path.Combine(filePath, "Highlighter", fileName);
            if (!File.Exists(fullPath))
                throw new FileNotFoundException($"Файл не найден: {fullPath}");

            using (var workbook = new XLWorkbook(fullPath))
            {
                if (!workbook.Worksheets.Contains(sheetName))
                    throw new ArgumentException($"Лист '{sheetName}' не найден в файле {fileName}");

                var ws = workbook.Worksheet(sheetName);
                var rows = ws.RangeUsed().RowsUsed(); // все используемые строки

                bool isFirst = true;
                foreach (var row in rows)
                {
                    if (isFirst)
                    {
                        // предполагаем, что первая строка — заголовки
                        isFirst = false;
                        continue;
                    }

                    string subsys = row.Cell(1).GetString().Trim();
                    if (string.IsNullOrEmpty(subsys))
                        continue;

                    var rec = new PhysData
                    {
                        PipeLine = subsys,
                        Tonnage = ReadDouble(row.Cell(2)),
                        Length = ReadDouble(row.Cell(3)),
                        Volume = ReadDouble(row.Cell(4))
                    };

                    result.Add(rec);
                }
            }

            return result;
        }

        public static List<PhysData> LoadFromCsv(string filePath, string fileName = "subsystems_summary_test.csv")
        {
            var result = new List<PhysData>();
            string fullPath = Path.Combine(filePath, "Highlighter", fileName);

            if (!File.Exists(fullPath))
                throw new FileNotFoundException($"Файл не найден: {fullPath}");

            using (var reader = new StreamReader(fullPath))
            {
                bool isFirst = true;
                while (!reader.EndOfStream)
                {
                    var line = reader.ReadLine();
                    if (isFirst)
                    {
                        isFirst = false; // Пропускаем заголовки
                        continue;
                    }

                    if (string.IsNullOrWhiteSpace(line))
                        continue;

                    var parts = line.Split(';');
                    if (parts.Length < 4)
                        continue;

                    string lines = parts[0].Trim();
                    if (string.IsNullOrEmpty(lines))
                        continue;

                    var rec = new PhysData
                    {
                        PipeLine = lines,
                        Length = ParseDouble(parts[1]),
                        Tonnage = ParseDouble(parts[2]),
                        Volume = ParseDouble(parts[3])
                    };

                    result.Add(rec);
                }
            }

            return result;
        }

        // Утилита для парсинга чисел с учетом возможной запятой
        private static double ParseDouble(string s)
        {
            return double.TryParse(s.Replace(',', '.'), NumberStyles.Any, CultureInfo.InvariantCulture, out double val)
                ? val
                : 0.0;
        }

        /// <summary>
        /// Преобразует object из ClosedXML в double, учитывая разные варианты хранения
        /// (число, строка с точкой или запятой).
        /// </summary>
        private static double ReadDouble(IXLCell cell)
        {
            if (cell == null || cell.IsEmpty())
                return 0;

            // 1) Если в ячейке действительно число — возвращаем его
            if (cell.DataType == XLDataType.Number)
                return cell.GetDouble();

            // 2) Иначе читаем как строку и парсим
            var s = cell.GetString().Trim();

            // Сначала через Invariant (точка), потом через текущую культуру (запятая)
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d) ||
                double.TryParse(s, NumberStyles.Any, CultureInfo.CurrentCulture, out d))
            {
                return d;
            }

            return 0;
        }
    }
}
