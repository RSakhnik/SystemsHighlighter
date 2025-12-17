using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using SystemsHighlighter.Tools;
using static SystemsHighlighter.Tools.SystemsData;

namespace SystemsHighlighter
{
    /// 
    /// Загрузчик данных из Excel-файла UPS_GUID.xlsx для одного листа (части модели).
    /// 
    public class SystemsDataLoader
    {
        private const string ExcelFileName = "UPS_GUID.xlsx";
        // Код части модели (имя листа), который загружается этим экземпляром
        private readonly string _partName;
        private string _filePath;

        // Результат: словарь код системы → SystemClass
        public Dictionary<string, SystemsData.SystemClass> _systems
            = new Dictionary<string, SystemsData.SystemClass>();

        /// <summary>
        /// Конструктор: принимает код части модели (имя листа) и сразу загружает данные.
        /// </summary>
        /// <param name="partName">Имя листа Excel, соответствующее части модели</param>
        //public SystemsDataLoader(string partName, string filePath)
        //{
        //    if (string.IsNullOrWhiteSpace(partName))
        //        throw new ArgumentException("Имя листа не может быть пустым", nameof(partName));
        //    _partName = partName.Trim();
        //    _filePath = filePath;
        //    Load();
        //}

        public SystemsDataLoader(string partName, string filePath, bool deferLoad)
        {
            if (string.IsNullOrWhiteSpace(partName))
                throw new ArgumentException("Имя листа не может быть пустым", nameof(partName));
            _partName = partName.Trim();
            _filePath = filePath;
            //LoadAsync(partName, filePath);
        }

        /// <summary>
        /// Возвращает словарь всех систем для указанной части модели.
        /// </summary>
        public IReadOnlyDictionary<string, SystemsData.SystemClass> Systems
        {
            get { return _systems; }
        }

        /*
        /// <summary>
        /// Выполняет чтение файла и парсинг только нужного листа.
        /// </summary>
        private void Load()
        {
            // Определяем папку сборки
            //string assemblyFolder = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            //if (assemblyFolder == null)
            //    throw new DirectoryNotFoundException("Не удалось определить папку сборки.");

            string filePath = Path.Combine(_filePath, "Highlighter", ExcelFileName);
            if (!File.Exists(filePath))
                throw new FileNotFoundException("Excel-файл не найден: " + filePath);

            // Открываем книгу
            using (var workbook = new XLWorkbook(filePath))
            {
                // Ищем лист с нужным именем
                var worksheet = workbook.Worksheets.FirstOrDefault(ws =>
                    string.Equals(ws.Name.Trim(), _partName, StringComparison.OrdinalIgnoreCase));

                if (worksheet == null)
                    throw new ArgumentException(
                        string.Format("Лист с именем '{0}' не найден в файле {1}.", _partName, filePath));

                // Читаем все строки, пропуская заголовок (первая строка)
                var rows = worksheet.RowsUsed().ToList();
                for (int i = 1; i < rows.Count; i++)
                {
                    var row = rows[i];
                    string section = row.Cell(1).GetString().Trim();
                    string systemCode = row.Cell(2).GetString().Trim();
                    string subsystemCode = row.Cell(3).GetString().Trim();
                    string guidText = row.Cell(4).GetString().Trim();

                    if (string.IsNullOrEmpty(systemCode) ||
                        string.IsNullOrEmpty(subsystemCode) ||
                        string.IsNullOrEmpty(guidText))
                    {
                        continue;
                    }

                    Guid elementGuid;
                    if (!Guid.TryParse(guidText, out elementGuid))
                        continue;

                    // Получаем или создаём систему
                    SystemsData.SystemClass systemClass;
                    if (!_systems.TryGetValue(systemCode, out systemClass))
                    {
                        systemClass = new SystemsData.SystemClass(systemCode);
                        _systems[systemCode] = systemClass;
                    }

                    // Получаем или создаём подсистему
                    var subsystem = systemClass.GetOrCreateSubsystem(subsystemCode);

                    // Добавляем элемент (Guid и Pipeline)
                    var element = new SystemsData.Element(elementGuid, section);
                    subsystem.AddElement(elementGuid, section);
                }
            }
        }

        public static async Task<SystemsDataLoader> LoadAsync(string partName, string filePath)
        {
            return await Task.Run(() =>
            {
                var loader = new SystemsDataLoader(partName, filePath, deferLoad: true);
                loader.Load();
                return loader;
            });
        }

        public static async Task<SystemsDataLoader> LoadMultiplePartsFromSingleFileAsync(IEnumerable<string> partNames, string filePath)
        {
            return await Task.Run(() =>
            {
                if (partNames == null || !partNames.Any())
                    throw new ArgumentException("Список частей модели пуст.", nameof(partNames));

                string fullPath = Path.Combine(filePath, "Highlighter", ExcelFileName);
                if (!File.Exists(fullPath))
                    throw new FileNotFoundException("Excel-файл не найден: " + fullPath);

                var mergedSystems = new Dictionary<string, SystemsData.SystemClass>();

                using (var workbook = new XLWorkbook(fullPath))
                {
                    foreach (var partName in partNames)
                    {
                        var worksheet = workbook.Worksheets.FirstOrDefault(ws =>
                            string.Equals(ws.Name.Trim(), partName.Trim(), StringComparison.OrdinalIgnoreCase));

                        if (worksheet == null)
                            throw new ArgumentException($"Лист с именем '{partName}' не найден в файле {fullPath}.");

                        var rows = worksheet.RowsUsed().ToList();
                        for (int i = 1; i < rows.Count; i++)
                        {
                            var row = rows[i];
                            string section = row.Cell(1).GetString().Trim();
                            string systemCode = row.Cell(2).GetString().Trim();
                            string subsystemCode = row.Cell(3).GetString().Trim();
                            string guidText = row.Cell(4).GetString().Trim();

                            if (string.IsNullOrEmpty(systemCode) ||
                                string.IsNullOrEmpty(subsystemCode) ||
                                string.IsNullOrEmpty(guidText))
                                continue;

                            if (!Guid.TryParse(guidText, out var elementGuid))
                                continue;

                            if (!mergedSystems.TryGetValue(systemCode, out var systemClass))
                            {
                                systemClass = new SystemsData.SystemClass(systemCode);
                                mergedSystems[systemCode] = systemClass;
                            }

                            var subsystem = systemClass.GetOrCreateSubsystem(subsystemCode);
                            subsystem.AddElement(elementGuid, section);
                        }
                    }
                }

                var result = new SystemsDataLoader("merged", filePath, deferLoad: true);
                foreach (var kvp in mergedSystems)
                {
                    result._systems[kvp.Key] = kvp.Value;
                }

                return result;
            });
        }
        */

        public static async Task<SystemsDataLoader> LoadBySectionsFromCsvAsync(
    IEnumerable<string> partNames,
    string filePath, Dictionary<string, Guid> SectionMappings)
        {
            if (partNames == null || !partNames.Any())
                throw new ArgumentException("Список секций пуст.", nameof(partNames));

            string fullPath = Path.Combine(filePath, "Highlighter", "consolidated.csv");
            if (!File.Exists(fullPath))
                throw new FileNotFoundException("CSV-файл не найден: " + fullPath);

            var sectionSet = new HashSet<string>(partNames, StringComparer.OrdinalIgnoreCase);
            var mergedSystems = new Dictionary<string, SystemsData.SystemClass>(StringComparer.OrdinalIgnoreCase);

            using (var reader = new StreamReader(fullPath))
            {
                string headerLine = await reader.ReadLineAsync();
                if (string.IsNullOrWhiteSpace(headerLine))
                    throw new InvalidDataException("CSV-файл не содержит заголовков.");

                var headers = headerLine.Split(';').Select(h => h.Trim()).ToArray();

                int idxSection = Array.IndexOf(headers, "Секция");
                int idxSystem = Array.IndexOf(headers, "Код системы");
                int idxSubsystem = Array.IndexOf(headers, "Код подсистемы");
                int idxPipeLine = Array.IndexOf(headers, "Линия");
                int idxWeight = Array.IndexOf(headers, "Масса, кг");
                int idxLength = Array.IndexOf(headers, "Длина, м");
                int idxVolume = Array.IndexOf(headers, "Объём, м3");
                int idxDiaInch = Array.IndexOf(headers, "Dia-inch");
                int idxGuid = Array.IndexOf(headers, "Guid");

                if (idxSection == -1 || idxSystem == -1 || idxSubsystem == -1 || idxGuid == -1)
                    throw new InvalidDataException("CSV-файл не содержит одну или несколько обязательных колонок: 'Секция', 'Код системы', 'Код подсистемы', 'Guid'.");

                while (!reader.EndOfStream)
                {
                    var line = await reader.ReadLineAsync();
                    if (string.IsNullOrWhiteSpace(line)) continue;

                    var cols = line.Split(';');
                    if (cols.Length != headers.Length) continue;

                    string section = cols[idxSection].Trim();
                    if (!sectionSet.Contains(section)) continue;

                    string systemCode = cols[idxSystem].Trim();
                    string subsystemCode = cols[idxSubsystem].Trim();
                    string pipeline = cols[idxPipeLine].Trim();
                    string guidText = cols[idxGuid].Trim();

                    // численные параметры всего элемента
                    double weight = ParseDoubleSafe(cols[idxWeight]);
                    double length = ParseDoubleSafe(cols[idxLength]);
                    double volume = ParseDoubleSafe(cols[idxVolume]);
                    double diainch = ParseDoubleSafe(cols[idxDiaInch]);

                    var geometry_guids = guidText.Split('|');

                    if (geometry_guids.Length == 0) continue;

                    

                    // делим показатели между геометриями
                    double partWeight = weight / (geometry_guids.Length - 1);
                    double partLength = length / (geometry_guids.Length - 1);
                    double partVolume = volume / (geometry_guids.Length - 1);
                    double partDiaInch = diainch / (geometry_guids.Length - 1);

                    foreach (var geometry in geometry_guids)
                    {
                        if (string.IsNullOrEmpty(systemCode) ||
                            string.IsNullOrEmpty(subsystemCode) ||
                            !Guid.TryParse(geometry, out var elementGuid))
                        {
                            continue;
                        }

                        if (!mergedSystems.TryGetValue(systemCode, out var systemClass))
                        {
                            systemClass = new SystemsData.SystemClass(systemCode);
                            mergedSystems[systemCode] = systemClass;
                        }

                        var subsystem = systemClass.GetOrCreateSubsystem(subsystemCode);
                        subsystem.GetOrCreatePipeLine(pipeline)
                                 .AddElement(
                                     elementGuid,
                                     section,
                                     partWeight.ToString(),
                                     partLength.ToString(),
                                     partVolume.ToString(),
                                     partDiaInch.ToString(),
                                     SectionMappings);
                    }
                }
            }

            var result = new SystemsDataLoader("merged", filePath, deferLoad: true);
            foreach (var kvp in mergedSystems)
            {
                result._systems[kvp.Key] = kvp.Value;
            }

            return result;
        }

        static double ParseDoubleSafe(string s)
        {
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var d))
                return d;
            return 0.0;
        }


    }

}