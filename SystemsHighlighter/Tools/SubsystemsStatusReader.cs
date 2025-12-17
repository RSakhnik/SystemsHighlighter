using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;

namespace SystemsHighlighter.Tools
{
    public class SubsystemsStatusReader
    {
        private const string ExcelFileName = "SubSystemsStatus.xlsx";

        // Новый файл с приоритетами по линиям
        private const string LinePrioritiesExcelFileName = "Приоритеты (по линиям).xlsx";

        // Новый файл с разбиением на подсистемы по линиям
        private const string SubsystemsByLinesExcelFileName = "Разбиение на подсистемы (по линиям).xlsx";

        private readonly string _folderPath;

        public SubsystemsStatusReader(string folderPath)
        {
            if (string.IsNullOrWhiteSpace(folderPath))
                throw new ArgumentException("Путь к папке с Excel не задан.", nameof(folderPath));

            _folderPath = folderPath;
        }

        /// <summary>
        /// Информация по линии: тест-пакет, приоритет, проценты готовности, факт завершения испытаний.
        /// </summary>
        public sealed class LineStatus
        {
            public string LineName { get; set; }
            public string TestPackageName { get; set; }
            public string Priority { get; set; }
            public double WeldingPercent { get; set; }
            public double NdtPercent { get; set; }
            public bool TestsCompleted { get; set; }
        }

        /// <summary>
        /// Стандартный способ чтения статусов линий из SubSystemsStatus.xlsx.
        /// Требует наличия колонки "Приоритет".
        /// Старое поведение сохранено.
        /// </summary>
        public Dictionary<string, LineStatus> ReadLineStatuses()
        {
            return ReadLineStatusesInternal(
                requirePriorityColumn: true,
                externalPriorities: null);
        }

        /// <summary>
        /// Альтернативный способ: читаем статусы из SubSystemsStatus.xlsx,
        /// приоритеты добираем из файла "Приоритеты (по линиям).xlsx".
        /// Колонки "Приоритет" в основном файле может не быть.
        /// При этом приоритет добавляется ко всем линиям из старой структуры,
        /// для которых он задан во внешнем файле.
        /// </summary>
        public Dictionary<string, LineStatus> ReadLineStatusesWithExternalPriorities()
        {
            var externalPriorities = TryReadLinePrioritiesFromWorkbook();

            // читаем статусы (и частично приоритеты) из основного файла
            var result = ReadLineStatusesInternal(
                requirePriorityColumn: false,
                externalPriorities: externalPriorities);

            // гарантируем, что для каждой линии из файла приоритетов есть запись в result
            // (по нормализованному имени), чтобы старые структуры могли взять приоритет
            foreach (var kv in externalPriorities)
            {
                var normLineName = NormalizeLineNameLocal(kv.Key);
                if (string.IsNullOrWhiteSpace(normLineName))
                    continue;

                if (!result.TryGetValue(normLineName, out var lineStatus))
                {
                    // Линии не было вообще — создаём пустую, но с приоритетом.
                    lineStatus = new LineStatus
                    {
                        LineName = normLineName,
                        TestPackageName = null,
                        Priority = kv.Value,
                        WeldingPercent = 0.0,
                        NdtPercent = 0.0,
                        TestsCompleted = false
                    };
                    result[normLineName] = lineStatus;
                }
                else
                {
                    // Линия была, но приоритета не было — допишем.
                    if (string.IsNullOrWhiteSpace(lineStatus.Priority))
                    {
                        lineStatus.Priority = kv.Value;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Общий внутренний метод чтения статусов линий.
        /// externalPriorities — словарь нормализованных имён линий -> приоритет (лейбл листа).
        /// </summary>
        private Dictionary<string, LineStatus> ReadLineStatusesInternal(
            bool requirePriorityColumn,
            Dictionary<string, string> externalPriorities)
        {
            var result = new Dictionary<string, LineStatus>(StringComparer.OrdinalIgnoreCase);

            var path = Path.Combine(_folderPath, ExcelFileName);
            if (!File.Exists(path))
                throw new FileNotFoundException($"Файл {ExcelFileName} не найден по пути: {path}");

            using (var wb = new XLWorkbook(path))
            {
                var ws = wb.Worksheet(1);
                var usedRange = ws.RangeUsed();
                if (usedRange == null)
                    return result;

                int firstRow = usedRange.FirstRow().RowNumber();
                int lastRow = usedRange.LastRow().RowNumber();
                int firstCol = usedRange.FirstColumn().ColumnNumber();
                int lastCol = usedRange.LastColumn().ColumnNumber();

                // Ищем номера нужных колонок по тексту заголовков в первых нескольких строках.
                int colPriority = FindColumnByHeader(ws, firstRow, lastRow, firstCol, lastCol, "приоритет");
                int colTpOrLine = FindColumnByHeader(ws, firstRow, lastRow, firstCol, lastCol, "тест-пакета/линия");
                int colWelding = FindColumnByHeader(ws, firstRow, lastRow, firstCol, lastCol, "процент готовности по сварке");
                int colNdt = FindColumnByHeader(ws, firstRow, lastRow, firstCol, lastCol, "процент готовности по нк");
                int colTestsCompleted = FindColumnByHeaderFromRight(ws, firstRow, lastRow, firstCol, lastCol, "отметка об испытаниях");

                if (colTpOrLine <= 0)
                    throw new InvalidOperationException("Не удалось найти колонку '№ тест-пакета/линия' в файле SubSystemsStatus.");

                // В старом режиме колонка приоритета обязательна, в новом – нет.
                if (requirePriorityColumn && colPriority <= 0)
                    throw new InvalidOperationException("Не удалось найти колонку 'Приоритет' в файле SubSystemsStatus.");

                // Будем считать, что данные начинаются после заголовков (берём первую строку, где в столбце TP/линия что-то есть, и идём от неё).
                int dataStartRow = DetectDataStartRow(ws, firstRow, lastRow, colTpOrLine);

                string currentTpName = null;
                double currentWeldingPercent = 0.0;
                double currentNdtPercent = 0.0;
                bool currentTestsCompleted = false;
                string currentPriorityForBlock = null;

                for (int row = dataStartRow; row <= lastRow; row++)
                {
                    string tpOrLineText = ws.Cell(row, colTpOrLine).GetString()?.Trim();
                    string tpStatus = colTestsCompleted > 0 ? ws.Cell(row, colTestsCompleted).GetString()?.Trim() : null;
                    string priorityText = colPriority > 0 ? ws.Cell(row, colPriority).GetString()?.Trim() : null;

                    if (string.IsNullOrWhiteSpace(tpOrLineText))
                        continue;

                    // Строка тест-пакета
                    if (IsTestPackageRow(tpOrLineText))
                    {
                        if (string.IsNullOrEmpty(tpStatus))
                            tpStatus = "Не принято";

                        currentTpName = tpOrLineText + " (" + tpStatus + ")";

                        // Обновляем "текущий приоритет" для блока TP.
                        if (!string.IsNullOrWhiteSpace(priorityText))
                            currentPriorityForBlock = priorityText;

                        // Проценты / флаг берём с этой строки, если колонки найдены.
                        currentWeldingPercent = colWelding > 0 ? ParsePercent(ws.Cell(row, colWelding)) : 0.0;
                        currentNdtPercent = colNdt > 0 ? ParsePercent(ws.Cell(row, colNdt)) : 0.0;
                        currentTestsCompleted = colTestsCompleted > 0 && ParseBool(ws.Cell(row, colTestsCompleted));

                        continue;
                    }

                    // Иначе это строка линии, но только если уже есть активный TP.
                    if (string.IsNullOrEmpty(currentTpName))
                        continue;

                    string lineNameRaw = tpOrLineText;
                    var normLineName = NormalizeLineNameLocal(lineNameRaw);
                    if (string.IsNullOrWhiteSpace(normLineName))
                        continue;

                    // Приоритет для линии:
                    // 1) собственный текст в строке;
                    // 2) внешний файл приоритетов по линиям;
                    // 3) приоритет блока TP.
                    string externalPriority = null;
                    if (externalPriorities != null && !string.IsNullOrWhiteSpace(normLineName))
                    {
                        externalPriorities.TryGetValue(normLineName, out externalPriority);
                    }

                    string linePriority =
                        !string.IsNullOrWhiteSpace(priorityText) ? priorityText :
                        !string.IsNullOrWhiteSpace(externalPriority) ? externalPriority :
                        currentPriorityForBlock;

                    if (!result.ContainsKey(normLineName))
                    {
                        result[normLineName] = new LineStatus
                        {
                            LineName = normLineName,
                            TestPackageName = currentTpName,
                            Priority = linePriority,
                            WeldingPercent = currentWeldingPercent,
                            NdtPercent = currentNdtPercent,
                            TestsCompleted = currentTestsCompleted
                        };
                    }
                    // Если одна и та же линия встречается несколько раз — берём первую.
                }
            }

            return result;
        }

        /// <summary>
        /// Читает файл "Приоритеты (по линиям).xlsx" и возвращает словарь:
        ///   ключ — НОРМАЛИЗОВАННОЕ имя линии,
        ///   значение — приоритет (имя листа, например "I Приоритет").
        /// Если файл не найден или в нём нет данных, возвращается пустой словарь.
        /// </summary>
        private Dictionary<string, string> TryReadLinePrioritiesFromWorkbook()
        {
            var result = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            var path = Path.Combine(_folderPath, LinePrioritiesExcelFileName);
            if (!File.Exists(path))
                return result;

            using (var wb = new XLWorkbook(path))
            {
                foreach (var ws in wb.Worksheets)
                {
                    string priorityLabel = ws.Name?.Trim();
                    if (string.IsNullOrWhiteSpace(priorityLabel))
                        continue;

                    var usedRange = ws.RangeUsed();
                    if (usedRange == null)
                        continue;

                    int firstRow = usedRange.FirstRow().RowNumber();
                    int lastRow = usedRange.LastRow().RowNumber();
                    int firstCol = usedRange.FirstColumn().ColumnNumber();
                    int lastCol = usedRange.LastColumn().ColumnNumber();

                    int colLine = FindColumnByHeader(ws, firstRow, lastRow, firstCol, lastCol, "трубопровод");
                    if (colLine <= 0)
                        continue;

                    // Находим первую строку с реальными данными (игнорируем заголовок "Трубопровод").
                    int dataStartRow = firstRow;
                    for (int row = firstRow; row <= lastRow; row++)
                    {
                        var txt = ws.Cell(row, colLine).GetString()?.Trim();
                        if (string.IsNullOrWhiteSpace(txt))
                            continue;

                        if (txt.ToLowerInvariant().Contains("трубопровод"))
                            continue;

                        dataStartRow = row;
                        break;
                    }

                    for (int row = dataStartRow; row <= lastRow; row++)
                    {
                        var lineNameRaw = ws.Cell(row, colLine).GetString()?.Trim();
                        var normLineName = NormalizeLineNameLocal(lineNameRaw);
                        if (string.IsNullOrWhiteSpace(normLineName))
                            continue;

                        if (!result.ContainsKey(normLineName))
                            result[normLineName] = priorityLabel;
                    }
                }
            }

            return result;
        }

        /// <summary>
        /// Читает файл "Разбиение на подсистемы (по линиям).xlsx" и возвращает
        /// структуру: ключ — имя подсистемы, значение — множество линий.
        /// </summary>
        public Dictionary<string, HashSet<string>> ReadSubsystemsStructureFromLines()
        {
            var result = new Dictionary<string, HashSet<string>>(StringComparer.OrdinalIgnoreCase);

            var path = Path.Combine(_folderPath, SubsystemsByLinesExcelFileName);
            if (!File.Exists(path))
                throw new FileNotFoundException($"Файл {SubsystemsByLinesExcelFileName} не найден по пути: {path}");

            using (var wb = new XLWorkbook(path))
            {
                var ws = wb.Worksheet(1);
                var usedRange = ws.RangeUsed();
                if (usedRange == null)
                    return result;

                int firstRow = usedRange.FirstRow().RowNumber();
                int lastRow = usedRange.LastRow().RowNumber();
                int firstCol = usedRange.FirstColumn().ColumnNumber();
                int lastCol = usedRange.LastColumn().ColumnNumber();

                int colSubsystem = FindColumnByHeader(ws, firstRow, lastRow, firstCol, lastCol, "подсистем");
                int colLine = FindColumnByHeader(ws, firstRow, lastRow, firstCol, lastCol, "линии");

                if (colSubsystem <= 0 || colLine <= 0)
                    throw new InvalidOperationException("Не удалось найти колонки 'Подсистемы' и 'Линии' в файле разбиения на подсистемы.");

                // Ищем первую строку с данными (ниже заголовков)
                int dataStartRow = firstRow;
                for (int row = firstRow; row <= lastRow; row++)
                {
                    var txtSub = ws.Cell(row, colSubsystem).GetString()?.Trim();
                    var txtLine = ws.Cell(row, colLine).GetString()?.Trim();

                    if (string.IsNullOrWhiteSpace(txtSub) && string.IsNullOrWhiteSpace(txtLine))
                        continue;

                    if ((txtSub?.ToLowerInvariant().Contains("подсистем") ?? false) ||
                        (txtLine?.ToLowerInvariant().Contains("линии") ?? false))
                        continue;

                    dataStartRow = row;
                    break;
                }

                for (int row = dataStartRow; row <= lastRow; row++)
                {
                    string subsysName = ws.Cell(row, colSubsystem).GetString()?.Trim();
                    string lineName = ws.Cell(row, colLine).GetString()?.Trim();

                    if (string.IsNullOrWhiteSpace(subsysName) || string.IsNullOrWhiteSpace(lineName))
                        continue;

                    if (!result.TryGetValue(subsysName, out var set))
                    {
                        set = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        result[subsysName] = set;
                    }

                    set.Add(lineName);
                }
            }

            return result;
        }

        // ===== вспомогательные методы =====

        private static string NormalizeLineNameLocal(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;
            return name.Trim();
        }

        private static bool IsTestPackageRow(string value)
        {
            if (string.IsNullOrWhiteSpace(value))
                return false;

            value = value.Trim();
            return value.StartsWith("TP", StringComparison.OrdinalIgnoreCase);
        }

        private static double ParsePercent(IXLCell cell)
        {
            var text = cell.GetString()?.Trim();
            if (string.IsNullOrWhiteSpace(text))
                return 0.0;

            text = text.Replace("%", "").Trim();
            text = text.Replace(',', '.');

            if (!double.TryParse(text, NumberStyles.Any, CultureInfo.InvariantCulture, out double val))
                return 0.0;

            if (val >= 0.0 && val <= 1.0)
                val *= 100.0;

            if (val < 0.0) val = 0.0;
            if (val > 100.0) val = 100.0;

            return val;
        }

        private static bool ParseBool(IXLCell cell)
        {
            var text = cell.GetString()?.Trim();
            if (string.IsNullOrWhiteSpace(text))
                return false;

            text = text.ToLowerInvariant();

            if (text == "да" || text == "yes" || text == "y" ||
                text == "completed" || text == "завершены" ||
                text == "true" || text == "1" || text.Contains("принят"))
                return true;

            if (bool.TryParse(text, out bool b))
                return b;

            return false;
        }

        private static int FindColumnByHeader(IXLWorksheet ws, int firstRow, int lastRow, int firstCol, int lastCol, string headerFragment)
        {
            if (string.IsNullOrWhiteSpace(headerFragment))
                return -1;

            headerFragment = headerFragment.ToLowerInvariant();
            int maxHeaderRow = Math.Min(firstRow + 5, lastRow); // ищем в первых нескольких строках

            for (int col = firstCol; col <= lastCol; col++)
            {
                for (int row = firstRow; row <= maxHeaderRow; row++)
                {
                    var txt = ws.Cell(row, col).GetString()?.Trim();
                    if (string.IsNullOrWhiteSpace(txt))
                        continue;

                    if (txt.ToLowerInvariant().Contains(headerFragment))
                        return col;
                }
            }

            return -1;
        }

        private static int FindColumnByHeaderFromRight(IXLWorksheet ws, int firstRow, int lastRow, int firstCol, int lastCol, string headerFragment)
        {
            if (string.IsNullOrWhiteSpace(headerFragment))
                return -1;

            headerFragment = headerFragment.ToLowerInvariant();
            int maxHeaderRow = Math.Min(firstRow + 5, lastRow);

            for (int col = lastCol; col >= firstCol; col--)
            {
                for (int row = firstRow; row <= maxHeaderRow; row++)
                {
                    var txt = ws.Cell(row, col).GetString()?.Trim();
                    if (string.IsNullOrWhiteSpace(txt))
                        continue;

                    if (txt.ToLowerInvariant().Contains(headerFragment))
                        return col;
                }
            }

            return -1;
        }

        private static int DetectDataStartRow(IXLWorksheet ws, int firstRow, int lastRow, int colTpOrLine)
        {
            for (int row = firstRow; row <= lastRow; row++)
            {
                var txt = ws.Cell(row, colTpOrLine).GetString()?.Trim();
                if (string.IsNullOrWhiteSpace(txt))
                    continue;

                // Пропускаем строку с самим заголовком.
                if (txt.Contains("тест-пакета/линия"))
                    continue;

                return row;
            }

            // На крайний случай возвращаем первую строку.
            return firstRow;
        }
    }

    public class SubsystemsStatusMapper
    {
        private readonly Dictionary<string, SystemsData.SystemClass> _systems;

        public SubsystemsStatusMapper(Dictionary<string, SystemsData.SystemClass> systems)
        {
            _systems = systems ?? throw new ArgumentNullException(nameof(systems));
        }

        public IReadOnlyDictionary<string, SystemsData.SystemClass> Systems => _systems;

        // ===== нормализация имён линий =====

        private static string NormalizeLineName(string name)
        {
            if (string.IsNullOrWhiteSpace(name))
                return null;

            return name.Trim();
        }

        /// <summary>
        /// Строит новый SubsystemsStatusMapper на основе альтернативного разбиения
        /// "подсистема → линии".
        ///
        /// Новая структура systems собирается так:
        /// - система берётся по линии из старого _systems (через BuildLineIndex);
        /// - подсистемы — по ключам из subsystemToLines;
        /// - в подсистемы добавляются трубопроводы с элементами из старой структуры.
        ///
        /// Старый _systems не трогаем, возвращаем новый мапер с новой структурой.
        /// </summary>
        public SubsystemsStatusMapper BuildRemappedMapperFromLineMapping(
            Dictionary<string, HashSet<string>> subsystemToLines)
        {
            if (subsystemToLines == null)
                throw new ArgumentNullException(nameof(subsystemToLines));

            // Индекс: нормализованное имя линии -> (система, элементы) из старой структуры
            var lineIndex = BuildLineIndex();

            // Новая структура systems:
            //   ключ — имя системы,
            //   значение — новый SystemClass с подсистемами по Excel-схеме
            var newSystems = new Dictionary<string, SystemsData.SystemClass>(StringComparer.OrdinalIgnoreCase);

            foreach (var subsEntry in subsystemToLines)
            {
                var subsysName = subsEntry.Key;
                if (string.IsNullOrWhiteSpace(subsysName))
                    continue;

                var lines = subsEntry.Value;
                if (lines == null || lines.Count == 0)
                    continue;

                foreach (var rawLineName in lines)
                {
                    if (string.IsNullOrWhiteSpace(rawLineName))
                        continue;

                    var normLineName = NormalizeLineName(rawLineName);
                    if (string.IsNullOrEmpty(normLineName))
                        continue;

                    // Жёстко: если такой линии нет в старой структуре — она нигде не используется
                    if (!lineIndex.TryGetValue(normLineName, out var lineInfo))
                        continue;

                    var systemName = lineInfo.SystemName;
                    if (string.IsNullOrWhiteSpace(systemName))
                        continue;

                    // --- система ---
                    if (!newSystems.TryGetValue(systemName, out var system))
                    {
                        system = new SystemsData.SystemClass(systemName)
                        {
                            Subsystems = new List<SystemsData.Subsystem>()
                        };
                        newSystems[systemName] = system;
                    }

                    // --- подсистема ---
                    var subsys = system.Subsystems
                        .FirstOrDefault(s => string.Equals(s.Name, subsysName, StringComparison.OrdinalIgnoreCase));

                    if (subsys == null)
                    {
                        subsys = new SystemsData.Subsystem(subsysName)
                        {
                            PipeLines = new List<SystemsData.PipeLine>()
                        };
                        system.Subsystems.Add(subsys);
                    }

                    // --- линия ---
                    var pipe = subsys.PipeLines
                        .FirstOrDefault(p => string.Equals(p.Name, normLineName, StringComparison.OrdinalIgnoreCase));

                    if (pipe == null)
                    {
                        pipe = new SystemsData.PipeLine(normLineName)
                        {
                            Elements = new List<SystemsData.Element>()
                        };
                        subsys.PipeLines.Add(pipe);
                    }

                    // --- элементы (GUID'ы и прочее) ---
                    if (lineInfo.Elements != null && lineInfo.Elements.Count > 0)
                        pipe.Elements.AddRange(lineInfo.Elements);
                }
            }

            // ==== CSV в буфер обмена: Подсистема;Линия ====

            var sb = new StringBuilder();
            sb.AppendLine("Подсистема;Линия");

            // Проходим по всем подсистемам и их линиям
            foreach (var system in newSystems.Values.OrderBy(s => s.Name, StringComparer.OrdinalIgnoreCase))
            {
                if (system.Subsystems == null)
                    continue;

                foreach (var subsys in system.Subsystems.OrderBy(s => s.Name, StringComparer.OrdinalIgnoreCase))
                {
                    if (subsys.PipeLines == null)
                        continue;

                    foreach (var pipe in subsys.PipeLines.OrderBy(p => p.Name, StringComparer.OrdinalIgnoreCase))
                    {
                        sb.Append(subsys.Name ?? string.Empty);
                        sb.Append(';');
                        sb.Append(pipe.Name ?? string.Empty);
                        sb.AppendLine();
                    }
                }
            }

            try
            {
                Clipboard.SetText(sb.ToString());
            }
            catch
            {
                // если поток не STA или ещё какой сюрприз — просто не копируем, но метод не падает
            }

            // В newSystems попадают только те линии,
            // которые одновременно есть и в экселе, и в старой структуре.
            return new SubsystemsStatusMapper(newSystems);
        }

        // ==== индекс по линиям ====

        private class LineIndexEntry
        {
            public string SystemName;
            public List<SystemsData.Element> Elements = new List<SystemsData.Element>();
        }

        /// <summary>
        /// Строит индекс по линиям на основе исходной структуры _systems:
        ///   линия -> (система, элементы).
        /// Используется при альтернативном разбиении подсистем по линиям.
        /// </summary>
        private Dictionary<string, LineIndexEntry> BuildLineIndex()
        {
            var index = new Dictionary<string, LineIndexEntry>(StringComparer.OrdinalIgnoreCase);

            foreach (var systemKv in _systems)
            {
                var systemName = systemKv.Key;
                var system = systemKv.Value;
                if (system == null || system.Subsystems == null)
                    continue;

                foreach (var subsystem in system.Subsystems)
                {
                    if (subsystem == null || subsystem.PipeLines == null)
                        continue;

                    foreach (var pipeLine in subsystem.PipeLines)
                    {
                        if (pipeLine == null || string.IsNullOrWhiteSpace(pipeLine.Name))
                            continue;

                        var normName = NormalizeLineName(pipeLine.Name);
                        if (string.IsNullOrEmpty(normName))
                            continue;

                        if (!index.TryGetValue(normName, out var entry))
                        {
                            entry = new LineIndexEntry
                            {
                                SystemName = systemName
                            };
                            index[normName] = entry;
                        }

                        if (pipeLine.Elements != null && pipeLine.Elements.Count > 0)
                            entry.Elements.AddRange(pipeLine.Elements);
                    }
                }
            }

            // Буфер обмена для отладки сейчас не используется.
            try
            {
                // если нужно будет снова копировать перечень линий, сюда можно вернуть код
            }
            catch
            {
                // тут можно залогировать, если очень хочется
            }

            return index;
        }

        /// <summary>
        /// Результат по подсистеме: текстовые поля + проценты + элементы (для Guid'ов).
        /// </summary>
        public class SubsystemStatusSummary
        {
            public string SystemName { get; set; }
            public string SubsystemName { get; set; }

            /// <summary>
            /// Все уникальные приоритеты, склеенные через \n.
            /// </summary>
            public string PrioritiesText { get; set; }

            /// <summary>
            /// Все уникальные номера тест-пакетов, склеенные через \n.
            /// </summary>
            public string TestPackagesText { get; set; }

            /// <summary>
            /// Средний процент готовности по сварке для подсистемы.
            /// </summary>
            public double AvgWeldingPercent { get; set; }

            /// <summary>
            /// Средний процент готовности по НК для подсистемы.
            /// </summary>
            public double AvgNdtPercent { get; set; }

            /// <summary>
            /// Процент строк "Испытания завершены = Да" внутри подсистемы.
            /// </summary>
            public double TestsCompletedPercent { get; set; }

            /// <summary>
            /// Все элементы подсистемы (через них можно вытащить Guid'ы).
            /// </summary>
            public List<SystemsData.Element> Elements { get; } = new List<SystemsData.Element>();
        }

        private class SubsystemAccum
        {
            public HashSet<string> Priorities { get; } =
                new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            public HashSet<string> TestPackages { get; } =
                new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            public double WeldingSum;
            public double NdtSum;
            public int LinesCount;
            public int CompletedCount;

            public List<SystemsData.Element> Elements { get; } = new List<SystemsData.Element>();
        }

        /// <summary>
        /// Исходный способ: строит агрегированные статусы по подсистемам,
        /// используя вложенные подсистемы из _systems.
        /// Даже если по линии нет статуса, но она есть в структуре,
        /// она попадает в подсистему с нулевыми значениями.
        /// </summary>
        public Dictionary<string, Dictionary<string, SubsystemStatusSummary>> BuildSubsystemStatuses(
            Dictionary<string, SubsystemsStatusReader.LineStatus> lineStatuses)
        {
            if (lineStatuses == null)
                throw new ArgumentNullException(nameof(lineStatuses));

            var accumDict = new Dictionary<string, Dictionary<string, SubsystemAccum>>(StringComparer.OrdinalIgnoreCase);

            foreach (var systemKv in _systems)
            {
                var systemName = systemKv.Key;
                var system = systemKv.Value;
                if (system == null || system.Subsystems == null || system.Subsystems.Count == 0)
                    continue;

                if (!accumDict.TryGetValue(systemName, out var subsystemsAccum))
                {
                    subsystemsAccum = new Dictionary<string, SubsystemAccum>(StringComparer.OrdinalIgnoreCase);
                    accumDict[systemName] = subsystemsAccum;
                }

                foreach (var subsystem in system.Subsystems)
                {
                    if (subsystem == null)
                        continue;

                    var subsysName = subsystem.Name;
                    if (string.IsNullOrWhiteSpace(subsysName))
                        continue;

                    if (!subsystemsAccum.TryGetValue(subsysName, out var accum))
                    {
                        accum = new SubsystemAccum();
                        subsystemsAccum[subsysName] = accum;
                    }

                    foreach (var pipeLine in subsystem.PipeLines)
                    {
                        if (pipeLine == null || string.IsNullOrWhiteSpace(pipeLine.Name))
                            continue;

                        var normName = NormalizeLineName(pipeLine.Name);
                        if (string.IsNullOrEmpty(normName))
                            continue;

                        // Если по линии нет статуса, создаём "пустой" статус с нулями,
                        // чтобы линия всё равно учитывалась в LinesCount.
                        if (!lineStatuses.TryGetValue(normName, out var lineStatus))
                        {
                            lineStatus = new SubsystemsStatusReader.LineStatus
                            {
                                LineName = normName,
                                TestPackageName = null,
                                Priority = null,
                                WeldingPercent = 0.0,
                                NdtPercent = 0.0,
                                TestsCompleted = false
                            };
                        }

                        if (!string.IsNullOrWhiteSpace(lineStatus.Priority))
                            accum.Priorities.Add(lineStatus.Priority.Trim());

                        if (!string.IsNullOrWhiteSpace(lineStatus.TestPackageName))
                            accum.TestPackages.Add(lineStatus.TestPackageName.Trim());

                        accum.WeldingSum += lineStatus.WeldingPercent;
                        accum.NdtSum += lineStatus.NdtPercent;
                        accum.LinesCount++;

                        if (lineStatus.TestsCompleted)
                            accum.CompletedCount++;

                        if (pipeLine.Elements != null && pipeLine.Elements.Count > 0)
                            accum.Elements.AddRange(pipeLine.Elements);
                    }
                }
            }

            return ConvertAccumsToSummaries(accumDict);
        }

        /// <summary>
        /// Новый способ: строит статусы подсистем по альтернативному разбиению
        /// "подсистема → линии", прочитанному из Excel.
        /// Элементы подсистемы вытаскиваются из старой структуры _systems по линиям.
        /// Даже если по линии нет статуса, но она есть в разбиении и в исходной структуре,
        /// она учитывается с нулевыми значениями.
        /// </summary>
        public Dictionary<string, Dictionary<string, SubsystemStatusSummary>> BuildSubsystemStatusesFromLineMapping(
            Dictionary<string, SubsystemsStatusReader.LineStatus> lineStatuses,
            Dictionary<string, HashSet<string>> subsystemToLines)
        {
            if (lineStatuses == null)
                throw new ArgumentNullException(nameof(lineStatuses));
            if (subsystemToLines == null)
                throw new ArgumentNullException(nameof(subsystemToLines));

            // Индекс: линия -> (система, элементы)
            var lineIndex = BuildLineIndex();

            var accumDict = new Dictionary<string, Dictionary<string, SubsystemAccum>>(StringComparer.OrdinalIgnoreCase);

            foreach (var subsEntry in subsystemToLines)
            {
                var subsysName = subsEntry.Key;
                if (string.IsNullOrWhiteSpace(subsysName))
                    continue;

                var lines = subsEntry.Value;
                if (lines == null || lines.Count == 0)
                    continue;

                foreach (var rawLineName in lines)
                {
                    if (string.IsNullOrWhiteSpace(rawLineName))
                        continue;

                    var normLineName = NormalizeLineName(rawLineName);
                    if (string.IsNullOrEmpty(normLineName))
                        continue;

                    // Если по линии нет статуса, создаём "пустой" статус с нулями.
                    if (!lineStatuses.TryGetValue(normLineName, out var lineStatus))
                    {
                        lineStatus = new SubsystemsStatusReader.LineStatus
                        {
                            LineName = normLineName,
                            TestPackageName = null,
                            Priority = null,
                            WeldingPercent = 0.0,
                            NdtPercent = 0.0,
                            TestsCompleted = false
                        };
                    }

                    if (!lineIndex.TryGetValue(normLineName, out var lineInfo))
                        continue; // по линии нет элементов/системы в старой структуре

                    var systemName = lineInfo.SystemName;
                    if (string.IsNullOrWhiteSpace(systemName))
                        continue;

                    if (!accumDict.TryGetValue(systemName, out var subsystemsAccum))
                    {
                        subsystemsAccum = new Dictionary<string, SubsystemAccum>(StringComparer.OrdinalIgnoreCase);
                        accumDict[systemName] = subsystemsAccum;
                    }

                    if (!subsystemsAccum.TryGetValue(subsysName, out var accum))
                    {
                        accum = new SubsystemAccum();
                        subsystemsAccum[subsysName] = accum;
                    }

                    // Приоритеты
                    if (!string.IsNullOrWhiteSpace(lineStatus.Priority))
                        accum.Priorities.Add(lineStatus.Priority.Trim());

                    // Тест-пакеты
                    if (!string.IsNullOrWhiteSpace(lineStatus.TestPackageName))
                        accum.TestPackages.Add(lineStatus.TestPackageName.Trim());

                    // Проценты
                    accum.WeldingSum += lineStatus.WeldingPercent;
                    accum.NdtSum += lineStatus.NdtPercent;
                    accum.LinesCount++;

                    // Испытания
                    if (lineStatus.TestsCompleted)
                        accum.CompletedCount++;

                    // Элементы (GUID'ы)
                    if (lineInfo.Elements != null && lineInfo.Elements.Count > 0)
                        accum.Elements.AddRange(lineInfo.Elements);
                }
            }

            return ConvertAccumsToSummaries(accumDict);
        }

        /// <summary>
        /// Общий метод конвертации аккумулированных данных в итоговые сводки.
        /// Используется и старым, и новым способом.
        /// </summary>
        private Dictionary<string, Dictionary<string, SubsystemStatusSummary>> ConvertAccumsToSummaries(
            Dictionary<string, Dictionary<string, SubsystemAccum>> accumDict)
        {
            var result = new Dictionary<string, Dictionary<string, SubsystemStatusSummary>>(StringComparer.OrdinalIgnoreCase);

            foreach (var sysKv in accumDict)
            {
                var systemName = sysKv.Key;
                var subsAccumDict = sysKv.Value;

                var subsSummaryDict = new Dictionary<string, SubsystemStatusSummary>(StringComparer.OrdinalIgnoreCase);
                result[systemName] = subsSummaryDict;

                foreach (var subsKv in subsAccumDict)
                {
                    var subsysName = subsKv.Key;
                    var acc = subsKv.Value;

                    double avgWelding = 0.0;
                    double avgNdt = 0.0;
                    double testsCompletedPercent = 0.0;

                    if (acc.LinesCount > 0)
                    {
                        avgWelding = acc.WeldingSum / acc.LinesCount;
                        avgNdt = acc.NdtSum / acc.LinesCount;
                        testsCompletedPercent = 100.0 * acc.CompletedCount / acc.LinesCount;
                    }

                    var summary = new SubsystemStatusSummary
                    {
                        SystemName = systemName,
                        SubsystemName = subsysName,
                        PrioritiesText = string.Join("\n", acc.Priorities.OrderBy(p => p)),
                        TestPackagesText = string.Join("\n", acc.TestPackages.OrderBy(tp => tp)),
                        AvgWeldingPercent = avgWelding,
                        AvgNdtPercent = avgNdt,
                        TestsCompletedPercent = testsCompletedPercent
                    };

                    summary.Elements.AddRange(acc.Elements);

                    subsSummaryDict[subsysName] = summary;
                }
            }

            return result;
        }

        /// <summary>
        /// Строит текстовый отчёт по подсистемам на основе BuildSubsystemStatuses().
        /// </summary>
        public static string BuildSubsystemsTextReport(
            Dictionary<string, Dictionary<string, SubsystemStatusSummary>> data)
        {
            if (data == null || data.Count == 0)
                return "Нет данных по подсистемам.";

            var sb = new StringBuilder();

            foreach (var sysKv in data.OrderBy(k => k.Key))
            {
                var systemName = sysKv.Key;
                sb.AppendLine("========================================");
                sb.AppendLine("Система: " + systemName);

                var subsDict = sysKv.Value;
                foreach (var subsKv in subsDict.OrderBy(k => k.Key))
                {
                    var s = subsKv.Value;

                    sb.AppendLine();
                    sb.AppendLine("  Подсистема: " + s.SubsystemName);

                    sb.AppendLine("    Приоритеты:");
                    if (!string.IsNullOrWhiteSpace(s.PrioritiesText))
                        sb.AppendLine("      " + s.PrioritiesText.Replace("\n", "\n      "));
                    else
                        sb.AppendLine("      (нет данных)");

                    sb.AppendLine("    Тест-пакеты:");
                    if (!string.IsNullOrWhiteSpace(s.TestPackagesText))
                        sb.AppendLine("      " + s.TestPackagesText.Replace("\n", "\n      "));
                    else
                        sb.AppendLine("      (нет данных)");

                    sb.AppendLine(string.Format(
                        CultureInfo.InvariantCulture,
                        "    Средняя готовность по сварке: {0:0.0}%\n" +
                        "    Средняя готовность по НК:     {1:0.0}%\n" +
                        "    Испытания завершены:          {2:0.0}% строк с 'Да'",
                        s.AvgWeldingPercent,
                        s.AvgNdtPercent,
                        s.TestsCompletedPercent));
                }

                sb.AppendLine();
            }

            return sb.ToString();
        }
    }
}
