using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Linq.Expressions;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using SystemsHighlighter.Tools;

namespace SystemsHighlighter.Tools
{
    public class TPData
    {
        public Dictionary<Guid, TechnicalInfo> Items = new Dictionary<Guid, TechnicalInfo>();

        private TPData(Dictionary<Guid, TechnicalInfo> items)
        {
            Items = items;
        }

        private static DateTime? ParseDateFromCell(IXLCell cell)
        {
            var str = cell.GetString().Trim();

            if (string.IsNullOrEmpty(str))
                return null;

            // Укажем точный формат: "день.месяц.год"
            if (DateTime.TryParseExact(str, "dd.MM.yyyy",
                System.Globalization.CultureInfo.InvariantCulture,
                System.Globalization.DateTimeStyles.None,
                out DateTime result))
            {
                return result;
            }

            return null;
        }

        private static double ParseDoubleFromCell(IXLCell cell)
        {
            var str = cell.GetString().Trim();

            if (string.IsNullOrEmpty(str))
                return 0.0;

            // Заменяем запятые на точки, убираем лишние символы (кроме цифр, точки, минуса)
            str = str.Replace(',', '.');

            var cleaned = new string(str.Where(c => char.IsDigit(c) || c == '.' || c == '-').ToArray());

            if (double.TryParse(cleaned, System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out double result))
            {
                return result;
            }

            return 0.0;
        }


        public class TechnicalInfo
        {
            public List<ForColorModelElementId> ModelElementIds { get; }
            public string Line { get; }
            public string Subsystem { get; }
            public string TPStatus { get; }
            public string TestType { get; }
            public double? Volume { get; }
            public double? TotalTpWdi { get; }
            public double? TotalTpLengthMeter { get; }
            public double? TotalTpWeightKg { get; }
            public DateTime? TestLimitsPreparationDate { get; }
            public DateTime? SubmittedToGazpromEng { get; }
            public DateTime? GazpromEngToUstayDate { get; }
            public string GazpromEngStatus { get; }
            public DateTime? ToUstayQcForPunch { get; }
            public DateTime? QcToTpOffice { get; }
            public DateTime? UstayQcRftAppDate { get; }
            public DateTime? GazpromQcAppDate { get; }
            public DateTime? TestDate { get; }
            public string Remarks { get; }
            public string PassportId { get; }
            public string Section {  get; }

            public TechnicalInfo(
                List<ForColorModelElementId> modelElementIds,
                string line,
                string subsystem,
                string tpStatus,
                string testType,
                double? volume,
                double? totalTpWdi,
                double? totalTpLengthMeter,
                double? totalTpWeightKg,
                DateTime? testLimitsPreparationDate,
                DateTime? submittedToGazpromEng,
                DateTime? gazpromEngToUstayDate,
                string gazpromEngStatus,
                DateTime? toUstayQcForPunch,
                DateTime? qcToTpOffice,
                DateTime? ustayQcRftAppDate,
                DateTime? gazpromQcAppDate,
                DateTime? testDate,
                string remarks,
                string passportId,
                string section)
            {
                ModelElementIds = modelElementIds;
                Line = line;
                Subsystem = subsystem;
                TPStatus = tpStatus;
                TestType = testType;
                Volume = volume;
                TotalTpWdi = totalTpWdi;
                TotalTpLengthMeter = totalTpLengthMeter;
                TotalTpWeightKg = totalTpWeightKg;
                TestLimitsPreparationDate = testLimitsPreparationDate;
                SubmittedToGazpromEng = submittedToGazpromEng;
                GazpromEngToUstayDate = gazpromEngToUstayDate;
                GazpromEngStatus = gazpromEngStatus;
                ToUstayQcForPunch = toUstayQcForPunch;
                QcToTpOffice = qcToTpOffice;
                UstayQcRftAppDate = ustayQcRftAppDate;
                GazpromQcAppDate = gazpromQcAppDate;
                TestDate = testDate;
                Remarks = remarks;
                PassportId = passportId;
                Section = section;
            }
        }

        public void ShowTPDataHeader()
        {
            int maxItems = 5;

            if (Items == null || Items.Count == 0)
            {
                MessageBox.Show("TPData пуст или не загружен.", "Ошибка");
                return;
            }

            var sb = new StringBuilder();
            sb.AppendLine("Пример записей TPData:");
            sb.AppendLine();

            // Заголовки
            sb.AppendLine("Line | TPStatus | TestType | TestPressure | TestDate");

            // Линия
            sb.AppendLine(new string('-', 70));

            // Выводим первые maxItems записей
            foreach (var kv in Items.Take(maxItems))
            {
                var info = kv.Value;

                sb.AppendLine($"{info.Line} | {info.TPStatus} | {info.TestType} | " +
                              $"{(info.Volume?.ToString("F1") ?? "-")} | " +
                              $"{(info.TestDate?.ToString("dd.MM.yyyy") ?? "-")}");
            }

            MessageBox.Show(sb.ToString(), $"TPData: первые {maxItems} записей");
        }

        public static async Task<TPData> LoadForSectionsAsyncCsv(IEnumerable<string> sectionCodes, string filePath, Dictionary<string, Guid> SectionMappings)
        {
            if (sectionCodes == null || !sectionCodes.Any())
                throw new ArgumentException("Список разделов (sectionCodes) пуст.");

            string fullPath = Path.Combine(filePath, "Highlighter", "consolidated.csv");
            if (!File.Exists(fullPath))
                throw new FileNotFoundException("Файл consolidated.csv не найден по пути: " + fullPath);

            var sectionSet = new HashSet<string>(sectionCodes.Select(s => s.Trim()), StringComparer.OrdinalIgnoreCase);
            var result = new Dictionary<Guid, TechnicalInfo>();

            using (var reader = new StreamReader(fullPath))
            {
                string headerLine = await reader.ReadLineAsync();
                if (string.IsNullOrWhiteSpace(headerLine))
                    throw new InvalidDataException("CSV-файл не содержит заголовков.");

                var headers = headerLine.Split(';');

                // Индексы нужных колонок
                int idxSection = Array.IndexOf(headers, "Секция");
                int idxGuid = Array.IndexOf(headers, "Guid");
                int idxLine = Array.IndexOf(headers, "Линия");
                int idxSubsystem = Array.IndexOf(headers, "Код подсистемы");
                int idxTpStatus = Array.IndexOf(headers, "TP STATUS");
                int idxTestType = Array.IndexOf(headers, "TEST TYPE");
                int idxVolume = Array.IndexOf(headers, "Объём, м3");
                int idxTotalWdi = Array.IndexOf(headers, "Dia-inch");
                int idxTotalLength = Array.IndexOf(headers, "Длина, м");
                int idxTotalWeight = Array.IndexOf(headers, "Масса, кг");
                int idxPrepDate = Array.FindIndex(headers, h => h.Contains("PREPARATION"));
                int idxSubmitted = Array.FindIndex(headers, h => h.Contains("SUBMITTED"));
                int idxEngToUstay = Array.FindIndex(headers, h => h.Contains("ENG TO USTAY"));
                //int idxEngStatus = -1; // Нет в consolidated.csv
                int idxToUstayQc = Array.FindIndex(headers, h => h.Contains("TO USTAY QC"));
                int idxQcToTpOffice = Array.FindIndex(headers, h => h.Contains("QC TO TP OFFICE"));
                int idxUstayQcApp = Array.FindIndex(headers, h => h.Contains("USTAY QC") && h.Contains("APP"));
                int idxGazpromQcApp = Array.FindIndex(headers, h => h.Contains("GAZPROM QC"));
                int idxTestDate = Array.IndexOf(headers, "TEST DATE");
                int idxRemarks = Array.IndexOf(headers, "REMARKS");
                int idxPassportId = Array.IndexOf(headers, "Паспорт");

                while (!reader.EndOfStream)
                {
                    string line = await reader.ReadLineAsync();
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var cols = line.Split(';');
                    if (cols.Length < headers.Length) continue;

                    string section = cols[idxSection].Trim();
                    if (!sectionSet.Contains(section)) continue;

                    var geometry_guids = cols[idxGuid].Split('|');

                    string guidStr = geometry_guids[0];
                    if (!Guid.TryParse(guidStr, out Guid elementGuid)) continue;

                    if (!SectionMappings.TryGetValue(section, out Guid modelPartId))
                        continue;

                    var guids = new List<ForColorModelElementId>();

                    foreach ( var guid in geometry_guids)
                    {
                        if (Guid.TryParse(guid, out Guid gg)) guids.Add(new ForColorModelElementId(gg, modelPartId));
                    }


                    try
                    {
                        string lineCode = cols[idxLine];
                        string subSys = cols[idxSubsystem];
                        string tpStatus = cols[idxTpStatus];
                        string testType = cols[idxTestType];

                        double? volume = ParseDouble(cols[idxVolume]);
                        double? totalWdi = ParseDouble(cols[idxTotalWdi]); // не используется
                        double? totalLength = ParseDouble(cols[idxTotalLength]);
                        double? totalWeight = ParseDouble(cols[idxTotalWeight]);

                        DateTime? prepDate = ParseDate(cols[idxPrepDate]);
                        DateTime? submittedGazprom = ParseDate(cols[idxSubmitted]);
                        DateTime? engToUstay = ParseDate(cols[idxEngToUstay]);

                        string engStatus = ""; // если нужно — можно парсить отдельно

                        DateTime? toUstayQc = ParseDate(cols[idxToUstayQc]);
                        DateTime? qcToTpOffice = ParseDate(cols[idxQcToTpOffice]);
                        DateTime? ustayQcApp = ParseDate(cols[idxUstayQcApp]);
                        DateTime? gazpromQcApp = ParseDate(cols[idxGazpromQcApp]);
                        DateTime? testDate = ParseDate(cols[idxTestDate]);

                        string remarks = idxRemarks != -1 ? cols[idxRemarks] : "";

                        string passportid = cols[idxPassportId];

                        var info = new TechnicalInfo(
                                                    guids, lineCode, subSys, tpStatus, testType,
                                                    volume, totalWdi, totalLength, totalWeight,
                                                    prepDate, submittedGazprom, engToUstay,
                                                    engStatus, toUstayQc, qcToTpOffice, ustayQcApp,
                                                    gazpromQcApp, testDate, remarks, passportid, section
                                                    );

                        result[elementGuid] = info;
                    }
                    catch (Exception ex)
                    {
                        var message = new StringBuilder();
                        message.AppendLine("Ошибка при разборе строки CSV.");
                        message.AppendLine($"Исключение: {ex.GetType().Name} - {ex.Message}");
                        message.AppendLine("Индексы и значения:");

                        void AddField(string name, int index)
                        {
                            string val = index >= 0 && index < cols.Length ? cols[index] : "<нет значения>";
                            message.AppendLine($"  {name} (index {index}): {val}");
                        }

                        AddField("Line", idxLine);
                        AddField("TP Status", idxTpStatus);
                        AddField("Test Type", idxTestType);
                        AddField("Volume", idxVolume);
                        AddField("Total WDI", idxTotalWdi);
                        AddField("Total Length", idxTotalLength);
                        AddField("Total Weight", idxTotalWeight);
                        AddField("Preparation Date", idxPrepDate);
                        AddField("Submitted to Gazprom", idxSubmitted);
                        AddField("Eng to Ustay", idxEngToUstay);
                        AddField("To Ustay QC", idxToUstayQc);
                        AddField("QC to TP Office", idxQcToTpOffice);
                        AddField("Ustay QC App", idxUstayQcApp);
                        AddField("Gazprom QC App", idxGazpromQcApp);
                        AddField("Test Date", idxTestDate);
                        AddField("Remarks", idxRemarks);

                        message.AppendLine("Исходная строка:");
                        message.AppendLine(string.Join(" | ", cols));

                        throw new FormatException(message.ToString(), ex);
                    }

                    
                }
            }

            return new TPData(result);
        }

        // Утилиты

        private static double? ParseDouble(string value)
        {
            if (double.TryParse(value.Replace(',', '.'), System.Globalization.NumberStyles.Any,
                System.Globalization.CultureInfo.InvariantCulture, out var result))
                return result;
            return null;
        }

        private static DateTime? ParseDate(string value)
        {
            if (string.IsNullOrWhiteSpace(value)) return null;

            string[] formats = new[] { "dd-MM-yyyy", "dd.MM.yyyy", "yyyy-MM-dd", "dd/MM/yyyy", "yyyy/MM/dd" };
            foreach (var fmt in formats)
            {
                if (DateTime.TryParseExact(value.Trim(), fmt, null,
                    System.Globalization.DateTimeStyles.None, out var dt))
                    return dt;
            }
            return null;
        }



    }


}
