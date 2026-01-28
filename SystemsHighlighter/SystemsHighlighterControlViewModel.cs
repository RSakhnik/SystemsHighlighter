using Ascon.Pilot.Bim;
using Ascon.Pilot.Bim.SDK;
using Ascon.Pilot.Bim.SDK.ModelTab.SidebarTab;
using Ascon.Pilot.Bim.SDK.ModelViewer;
using Ascon.Pilot.Bim.SDK.Search;
using Ascon.Pilot.SDK;
using Ascon.Pilot.SDK.Data;
using ClosedXML.Excel;
using DocumentFormat.OpenXml.Bibliography;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using Microsoft.Win32;
using Prism.Commands;
using QuestPDF.Drawing;
using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.IO;
using System.IO.Ports;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;
using System.Windows.Media;
using SystemsHighlighter.Tools;
using static System.Net.WebRequestMethods;
using static SystemsHighlighter.SystemsHighlighterControlViewModel;
using static SystemsHighlighter.Tools.SubsystemsStatusReader;
using static SystemsHighlighter.Tools.SystemsData;
using static SystemsHighlighter.Tools.TPData;
using ColorCircle = System.Windows.Media.Color;
using IContainer = QuestPDF.Infrastructure.IContainer;

namespace SystemsHighlighter
{
    public class SystemsHighlighterControlViewModel : INotifyPropertyChanged
    {
        #region Контейнеры
        string _filePath;
        string ProjectName = "УПС";
        
        //Коллекция узлов дерева для привязки в XAML
        private readonly ObservableCollection<SystemNodeViewModel> _treeNodes = new ObservableCollection<SystemNodeViewModel>();
        public ObservableCollection<SystemNodeViewModel> TreeNodes => _treeNodes;

        public ObservableCollection<TP_LegendItem> TP_Legend { get; } = new ObservableCollection<TP_LegendItem>();

        public ObservableCollection<PP_LegendItem> PP_Legend { get; } = new ObservableCollection<PP_LegendItem>();
        public ObservableCollection<PP_PassportLegendNode> PP_LegendTree { get; } = new ObservableCollection<PP_PassportLegendNode>();



        private IModelViewer _modelViewer;
        private Dictionary<string, SystemClass> Systems;

        private List<string> LoadedModelParts = new List<string>();

        private TPData TestData;

        /// <summary>
        /// Все «занятые» цвета для узлов.
        /// </summary>
        private readonly List<ColorCircle> _usedColors = new List<ColorCircle>();
        private readonly Random _rnd = new Random();


        // Заранее загружаем все физданные
        public List<PhysData> _allPhysData;

        private double _totalTonnage;
        public double TotalTonnage
        {
            get => _totalTonnage;
            private set { _totalTonnage = value; OnPropertyChanged(); }
        }

        private double _totalLength;
        public double TotalLength
        {
            get => _totalLength;
            private set { _totalLength = value; OnPropertyChanged(); }
        }

        private double _totalVolume;
        public double TotalVolume
        {
            get => _totalVolume;
            private set { _totalVolume = value; OnPropertyChanged(); }
        }

        private bool _shouldRecalcPhys;
        public bool ShouldRecalcPhys
        {
            get => _shouldRecalcPhys;
            set
            {
                _shouldRecalcPhys = value;
                OnPropertyChanged();
            }
        }


        private int _loadProgress;
        private string _loadStatus;
        private bool _isLoading;

        public int LoadProgress
        {
            get => _loadProgress;
            set
            {
                if (_loadProgress != value)
                {
                    _loadProgress = value;
                    OnPropertyChanged(nameof(LoadProgress));
                }
            }
        }

        public string LoadStatus
        {
            get => _loadStatus;
            set
            {
                if (_loadStatus != value)
                {
                    _loadStatus = value;
                    OnPropertyChanged(nameof(LoadStatus));
                }
            }
        }

        public bool IsLoading
        {
            get => _isLoading;
            set
            {
                if (_isLoading != value)
                {
                    _isLoading = value;
                    OnPropertyChanged(nameof(IsLoading));
                }
            }
        }

        public bool HideUnhighlighted_mode = false;

        private Dictionary<Guid, string> partMap_guid_string = new Dictionary<Guid, string>();

        private Dictionary<string, Guid> partMap_string_guid = new Dictionary<string, Guid>();

        #endregion

        #region Инициализация элементов управления
        public SystemsHighlighterControlViewModel()
        {
            SearchCommand = new DelegateCommand(OnSearch);
            ClearSearchCommand = new DelegateCommand(OnClearSearch);
            SelectAllTopLevelCommand = new DelegateCommand(OnSelectAllTopLevel);
            ClearAllCommand = new DelegateCommand(OnClearAll);
            TP_ShowStatusCommand = new DelegateCommand(TP_ShowStatus);
            TP_ShowTypeCommand = new DelegateCommand(TP_ShowType);

            SubsystemsSearchCommand = new DelegateCommand(OnSubsystemsSearch);
            ClearSubsystemsSearchCommand = new DelegateCommand(OnClearSubsystemsSearch);

            ShowSubsystemsWeldingCommand = new DelegateCommand(OnShowSubsystemsWelding);
            ShowSubsystemsNdtCommand = new DelegateCommand(OnShowSubsystemsNdt);
            ShowSubsystemsTestsCommand = new DelegateCommand(OnShowSubsystemsTests);

            HideUnhighlightedSubsystemsCommand = new DelegateCommand(OnHideUnhighlightedSubsystems);
            RestoreVisibilitySubsystemsCommand = new DelegateCommand(OnRestoreVisibilitySubsystems);

            ActivateAllSubsystemsCommand = new DelegateCommand(OnActivateAllSubsystems);
            DeactivateAllSubsystemsCommand = new DelegateCommand(OnDeactivateAllSubsystems);



            if (_filePath == null);

            
            //_filePath = filePath;
        }

        public ISidebarTab SidebarTab { get; set; }

        public IModelViewer ModelViewer
        {
            get => _modelViewer;
            set
            {
                if (_modelViewer == value) return;

                // Отписываемся от старого
                if (_modelViewer != null)
                    _modelViewer.SelectionChanged -= OnSelectionChanged;

                _modelViewer = value;

                // Подписываемся на новое событие
                if (_modelViewer != null)
                    _modelViewer.SelectionChanged += OnSelectionChanged;

                OnPropertyChanged();
                RaiseAllCanExecuteChanged();
            }
        }

        public void OnModelLoaded(IModelViewer modelViewer)
        {
            ModelViewer = modelViewer;
            SidebarTab.IsVisible = ModelViewer != null;

            // PARTMAP
            var partMap = partMap_guid_string;

            //var partMap = new Dictionary<Guid, string>
            //    {
            //        { Guid.Parse("1d65306d-d5e3-492e-bb13-9e3327e0d67c"), "900" },
            //        { Guid.Parse("eaad9dd1-96ad-4ba8-9b45-d3abae092a89"), "200" },
            //        { Guid.Parse("138078c3-7b42-4555-91bc-d6c1211107b4"), "010" },
            //        { Guid.Parse("389deea2-ed77-413f-b110-57859325af38"), "300" },
            //        { Guid.Parse("3fb1ab10-3a28-4768-b7b5-a134053217f6"), "310" },
            //        { Guid.Parse("6c6abc95-a485-4aa0-8976-77ea17aac71b"), "320" },
            //        { Guid.Parse("254abf4e-acb5-461d-9339-0b3110e3eb60"), "330" },
            //        { Guid.Parse("29eb450d-4722-41bd-a60a-8dd8ddf0edf2"), "340" },
            //        { Guid.Parse("92e14574-3cf5-4bc0-af26-1c7762295b61"), "410" },
            //        { Guid.Parse("57f2df13-e310-4584-9adc-1babbce5057c"), "420" },
            //        { Guid.Parse("62b19655-c26b-4b28-92b7-c5568844e845"), "500" },
            //        { Guid.Parse("49107e4d-7f91-4dcc-9230-082d6c4af67a"), "720" },
            //        { Guid.Parse("971e43f1-a839-4e22-91c3-78240ed48e40"), "730" },
            //        { Guid.Parse("29c44c09-e9c1-414f-b3ae-07dbb271b53b"), "801" },
            //        { Guid.Parse("caf32630-af08-4ca9-8cc7-76807df05aa9"), "802" },
            //        { Guid.Parse("253af425-1622-4771-85d0-5f74fd9bf06f"), "803" },
            //        { Guid.Parse("7b256da3-fb91-4be6-8836-26b01d709b33"), "804" }
            //    };


            LoadedModelParts.Clear();
            foreach(var key in partMap.Keys)
            {
                if(_modelViewer.IsModelPartLoaded(key))
                {
                    LoadedModelParts.Add(partMap[key]);
                }
            }

            SelectedModelPartsText = null;
            SelectedModelPartsText = string.Join(Environment.NewLine, LoadedModelParts);
        }

        private void OnSelectionChanged(object sender, EventArgs e)
        {
            RaiseAllCanExecuteChanged();
        }

        private void RaiseAllCanExecuteChanged()
        {

        }
        #endregion

        #region Ручная загрузка части модели

        private ObservableCollection<string> _cmbModelPartsToLoadItems;
        public ObservableCollection<string> cmbModelPartsToLoadItems
        {
            get => _cmbModelPartsToLoadItems;
            set { _cmbModelPartsToLoadItems = value; OnPropertyChanged(); }
        }

        private string _cmbModelPartsToLoadSelected;
        public string cmbModelPartsToLoadSelected
        {
            get => _cmbModelPartsToLoadSelected;
            set { _cmbModelPartsToLoadSelected = value; OnPropertyChanged(); }
        }

        #endregion

        #region Инициализация данных
        /*public void InitModelData()
        {
            Systems = new Dictionary<string, SystemClass>();


            // 1) Собираем уникальные GUID-ы частей из видимых элементов
            var visiblePartGuids = ModelViewer
                .GetVisibleElements()
                .Select(elem => elem.ModelPartId)
                .Distinct()
                .ToList();

            if (!visiblePartGuids.Any())
            {
                // Нет видимых частей — выходим
                LoadedModelParts = new List<string>();
                return;
            }

            // 2) Словарь: GUID части → её код
            var partMap = new Dictionary<Guid, string>
                {
                    { Guid.Parse("138078c3-7b42-4555-91bc-d6c1211107b4"), "010" },
                    { Guid.Parse("eaad9dd1-96ad-4ba8-9b45-d3abae092a89"), "200" },
                    { Guid.Parse("389deea2-ed77-413f-b110-57859325af38"), "300" },
                    { Guid.Parse("00fbf2ff-9ffa-43d7-9325-c89be744fbab"), "310" },
                    { Guid.Parse("68dd413e-3af7-4ba3-9d47-1ac5d4820035"), "320" },
                    { Guid.Parse("ec8cbfde-cdc0-46a9-9230-ee7cc9a16130"), "330" },
                    { Guid.Parse("fac2815e-3798-431c-94f9-8a3092b8828e"), "340" },
                    { Guid.Parse("fb93465b-0724-49c5-ab58-3152161cf6ec"), "410" },
                    { Guid.Parse("7f537cbb-0fb9-4ada-b84c-fccc33147e7b"), "420" },
                    { Guid.Parse("152e2984-0407-4bd7-8734-4f5602e0dcdc"), "500" },
                    { Guid.Parse("66c60d6a-01b6-456f-929b-09eff70443bb"), "720" },
                    { Guid.Parse("25803b94-9282-4897-b073-b68bcfb8ac21"), "730" },
                    { Guid.Parse("15c1863b-e3b3-455b-9f6d-9f9cdc1adf78"), "801" },
                    { Guid.Parse("825f7973-8281-4986-9048-079c73c9a22a"), "802" },
                    { Guid.Parse("72c8ae4f-1547-40a6-ac3f-0d78d959998f"), "803" },
                    { Guid.Parse("594e6e20-f6b0-44b0-8bf5-462ec68c867a"), "804" },
                    { Guid.Parse("d5059009-1c00-4dac-8576-64c7ff55f0a5"), "900" },
                    { Guid.Parse("477b03fd-ca7e-4dc9-9a5f-4812c32d8485"), "910" },
                };

            // 3) Фильтруем только те GUID-ы, что и видимы, и есть в карте
            LoadedModelParts = visiblePartGuids
                .Where(guid => partMap.ContainsKey(guid))    // только загруженные
                .Select(guid => partMap[guid])                            // преобразуем в код
                .ToList();

            var part_systems = new SystemsDataLoader(LoadedModelParts.LastOrDefault(), _filePath);

            foreach (var system in part_systems.Systems)
            {
                if (Systems.ContainsKey(system.Key))
                    throw new InvalidOperationException($"Система {system.Key} уже была загружена для другой части.");
                Systems.Add(system.Key, system.Value);

            }

            BuildTree();

        }

        public void InitPhysData()
        {
            _allPhysData = Tools.PhysData.LoadFromExcel(_filePath);
        }*/

        public async Task InitModelDataAsync()
        {
            Systems = new Dictionary<string, SystemClass>();

            // PARTMAP
            //var partMap = new Dictionary<Guid, string>
            //    {
            //        { Guid.Parse("1d65306d-d5e3-492e-bb13-9e3327e0d67c"), "900" },
            //        { Guid.Parse("eaad9dd1-96ad-4ba8-9b45-d3abae092a89"), "200" },
            //        { Guid.Parse("138078c3-7b42-4555-91bc-d6c1211107b4"), "010" },
            //        { Guid.Parse("389deea2-ed77-413f-b110-57859325af38"), "300" },
            //        { Guid.Parse("3fb1ab10-3a28-4768-b7b5-a134053217f6"), "310" },
            //        { Guid.Parse("6c6abc95-a485-4aa0-8976-77ea17aac71b"), "320" },
            //        { Guid.Parse("254abf4e-acb5-461d-9339-0b3110e3eb60"), "330" },
            //        { Guid.Parse("29eb450d-4722-41bd-a60a-8dd8ddf0edf2"), "340" },
            //        { Guid.Parse("92e14574-3cf5-4bc0-af26-1c7762295b61"), "410" },
            //        { Guid.Parse("57f2df13-e310-4584-9adc-1babbce5057c"), "420" },
            //        { Guid.Parse("62b19655-c26b-4b28-92b7-c5568844e845"), "500" },
            //        { Guid.Parse("49107e4d-7f91-4dcc-9230-082d6c4af67a"), "720" },
            //        { Guid.Parse("971e43f1-a839-4e22-91c3-78240ed48e40"), "730" },
            //        { Guid.Parse("29c44c09-e9c1-414f-b3ae-07dbb271b53b"), "801" },
            //        { Guid.Parse("caf32630-af08-4ca9-8cc7-76807df05aa9"), "802" },
            //        { Guid.Parse("253af425-1622-4771-85d0-5f74fd9bf06f"), "803" },
            //        { Guid.Parse("7b256da3-fb91-4be6-8836-26b01d709b33"), "804" }
            //    };

            var part_map = partMap_guid_string;


            var partCode = LoadedModelParts.LastOrDefault();

            /*
            var systems = await Task.Run(() =>
            {
                var loader = new SystemsDataLoader(partCode, _filePath);
                return loader.Systems;
            });

            foreach (var system in systems)
            {
                if (Systems.ContainsKey(system.Key))
                {
                    Systems[system.Key].MergeFrom(system.Value);
                }
                else
                {
                    Systems.Add(system.Key, system.Value);
                }
            }*/

            var loader = await SystemsDataLoader.LoadBySectionsFromCsvAsync(LoadedModelParts, _filePath, partMap_string_guid);

            // доступ к объединённым системам
            var systems = loader.Systems;

            foreach (var system in systems)
            {
                if (Systems.ContainsKey(system.Key))
                {
                    Systems[system.Key].MergeFrom(system.Value);
                }
                else
                {
                    Systems.Add(system.Key, system.Value);
                }
            }

            await Application.Current.Dispatcher.InvokeAsync(() =>
            {
                BuildTree();
            });
        }

        public async Task InitPhysDataAsync()
        {
            //_allPhysData = await Task.Run(() => Tools.PhysData.LoadFromCsv(_filePath));
            await InitSubsystemStatusesAsync();
        }

        public async Task InitTPDataAsync()
        {
            TestData = await TPData.LoadForSectionsAsyncCsv(LoadedModelParts, _filePath, partMap_string_guid);
            BuildDateFieldOptions();
            PP_MakeLegend();
        }

        #endregion

        #region Дерево систем
        /// <summary>
        /// Строит TreeNodes из словаря Systems
        /// </summary>
        private void BuildTree()
        {
            _treeNodes.Clear();

            foreach (var sys in Systems.Values)
            {
                // Узел системы
                var systemElementIds = sys.Subsystems
                    .SelectMany(s => s.PipeLines)
                    .SelectMany(p => p.Elements)
                    .Select(e => (IModelElementId)e.ColorId)
                    .ToArray();

                var sysNode = new SystemNodeViewModel(sys.Name, systemElementIds, this);

                // Копим сумму по системе из сумм подсистем
                PhysicalTotals sysSum = default;

                foreach (var subs in sys.Subsystems)
                {
                    // Узел подсистемы
                    var subElementIds = subs.PipeLines
                        .SelectMany(p => p.Elements)
                        .Select(e => (IModelElementId)e.ColorId)
                        .ToArray();

                    var subNode = new SystemNodeViewModel(subs.Name, subElementIds, this);

                    // Копим сумму по подсистеме из сумм трубопроводов
                    PhysicalTotals subSum = default;

                    foreach (var pipe in subs.PipeLines)
                    {
                        var pipeElementIds = pipe.Elements
                            .Select(e => (IModelElementId)e.ColorId)
                            .ToArray();

                        // Лист: сумма по элементам трубопровода
                        PhysicalTotals pipeSum = SumElements(pipe.Elements);

                        var pipeNode = new SystemNodeViewModel(pipe.Name, pipeElementIds, this);
                        pipeNode.SetTotals(pipeSum);

                        subNode.Children.Add(pipeNode);
                        subSum += pipeSum;
                    }

                    subNode.SetTotals(subSum);
                    sysNode.Children.Add(subNode);
                    sysSum += subSum;
                }

                sysNode.SetTotals(sysSum);
                _treeNodes.Add(sysNode);
            }
        }


        public readonly struct PhysicalTotals
        {
            public decimal Weight { get; }
            public decimal Length { get; }
            public decimal Volume { get; }
            public decimal DiaInch { get; }

            public PhysicalTotals(decimal weight, decimal length, decimal volume, decimal diaInch)
            {
                Weight = weight; Length = length; Volume = volume; DiaInch = diaInch;
            }

            public static PhysicalTotals operator +(PhysicalTotals a, PhysicalTotals b) =>
                new PhysicalTotals(a.Weight + b.Weight, a.Length + b.Length, a.Volume + b.Volume, a.DiaInch + b.DiaInch);
        }

        private static decimal ParseDecimal(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return 0m;
            // Убираем пробелы-разделители тысяч и нормализуем разделитель дробной части:
            s = s.Trim().Replace(" ", "").Replace(",", ".");
            return decimal.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out var v) ? v : 0m;
        }

        private static PhysicalTotals SumElements(IEnumerable<Element> elements)
        {
            decimal w = 0, l = 0, v = 0, di = 0;
            foreach (var e in elements)
            {
                w += ParseDecimal(e.Weight) / 1000;
                l += ParseDecimal(e.Lenght);   // имя свойства в Element — "Lenght"
                v += ParseDecimal(e.Volume);
                di += ParseDecimal(e.DiaInch);
            }
            return new PhysicalTotals(w, l, v, di);
        }

        public static string Fmt(decimal value, int digits = 2) =>
            value.ToString($"N{digits}", CultureInfo.CurrentCulture);


        /// <summary>
        /// Поиск всех узлов (систем и подсистем) в дереве.
        /// </summary>
        private IEnumerable<SystemNodeViewModel> GetAllNodes()
        {
            foreach (var root in TreeNodes)
            {
                yield return root;
                foreach (var desc in GetDescendants(root))
                    yield return desc;
            }
        }

        private IEnumerable<SystemNodeViewModel> GetDescendants(SystemNodeViewModel node)
        {
            foreach (var ch in node.Children)
            {
                yield return ch;
                foreach (var sub in GetDescendants(ch))
                    yield return sub;
            }
        }


        #endregion

        #region Поиск по древу

        private readonly Dictionary<string, ColorCircle> _savedHighlightState = new Dictionary<string, ColorCircle>();

        private string _searchText;
        public string SearchText
        {
            get => _searchText;
            set
            {
                if (_searchText == value) return;
                _searchText = value;
                OnPropertyChanged();
            }
        }

        public DelegateCommand SearchCommand { get; }
        public DelegateCommand ClearSearchCommand { get; }

        private void OnSearch()
        {
            ApplyFilter(SearchText);
        }

        private void OnClearSearch()
        {
            SearchText = string.Empty;
            ApplyFilter(null);
        }

        /// <summary>
        /// Фильтрует дерево по одному или нескольким поисковым словам.
        /// Если filter начинается с '*', далее идёт список паттернов, разделённых ';'.
        /// </summary>
        public void ApplyFilter(string filter)
        {
            // 1) Сохраняем текущее состояние
            SaveHighlightState();

            _treeNodes.Clear();

            if (string.IsNullOrWhiteSpace(filter))
            {
                BuildTree();
            }
            else
            {
                bool multi = filter.StartsWith("*");
                List<string> patterns;

                if (multi)
                {
                    patterns = filter
                        .Substring(1)
                        .Split(new[] { ';' }, StringSplitOptions.RemoveEmptyEntries)
                        .Select(s => s.Trim())
                        .Where(s => s.Length > 0)
                        .ToList();

                    if (!patterns.Any())
                    {
                        BuildTree();
                        RestoreHighlightState();
                        RecalculatePhysicalSums();
                        return;
                    }
                }
                else
                {
                    patterns = new List<string> { filter.Trim() };
                }

                foreach (var sys in Systems.Values)
                {
                    bool sysMatches = patterns.Any(p =>
                        sys.Name.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0);

                    // Ищем подсистемы, которые совпали
                    var matchingSubsystems = sys.Subsystems
                        .Where(sub => patterns.Any(p =>
                            sub.Name.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0))
                        .ToList();

                    // Ищем совпадения в трубопроводах
                    foreach (var sub in sys.Subsystems)
                    {
                        var matchingPipes = sub.PipeLines
                            .Where(pipe => patterns.Any(p =>
                                pipe.Name.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0))
                            .ToList();

                        if (matchingPipes.Any() && !matchingSubsystems.Contains(sub))
                            matchingSubsystems.Add(sub);
                    }

                    if (sysMatches)
                    {
                        // Если совпала система — добавляем весь узел
                        _treeNodes.Add(CreateNode(sys)); // CreateNode должен строить Subsystem → PipeLine → Element
                    }
                    else if (matchingSubsystems.Any())
                    {
                        // Собираем все элементы этих подсистем (через их трубопроводы)
                        var allSystemElementIds = matchingSubsystems
                            .SelectMany(s => s.PipeLines)
                            .SelectMany(p => p.Elements)
                            .Select(e => (IModelElementId)e.ColorId)
                            .ToArray();

                        var sysNode = new SystemNodeViewModel(sys.Name, allSystemElementIds, this);

                        foreach (var subs in matchingSubsystems)
                        {
                            var subElementIds = subs.PipeLines
                                .SelectMany(p => p.Elements)
                                .Select(e => (IModelElementId)e.ColorId)
                                .ToArray();

                            var subNode = new SystemNodeViewModel(subs.Name, subElementIds, this);

                            // Если подсистема совпала — добавляем все её трубопроводы,
                            // если нет — только те, что совпали с фильтром
                            var pipesToAdd = patterns.Any(p =>
                                subs.Name.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0)
                                ? subs.PipeLines
                                : subs.PipeLines.Where(pipe =>
                                    patterns.Any(p =>
                                        pipe.Name.IndexOf(p, StringComparison.OrdinalIgnoreCase) >= 0));

                            foreach (var pipe in pipesToAdd)
                            {
                                var pipeElementIds = pipe.Elements
                                    .Select(e => (IModelElementId)e.ColorId)
                                    .ToArray();

                                var pipeNode = new SystemNodeViewModel(pipe.Name, pipeElementIds, this);
                                subNode.Children.Add(pipeNode);
                            }

                            sysNode.Children.Add(subNode);
                        }

                        _treeNodes.Add(sysNode);
                    }
                }
            }

            // 3) Восстанавливаем подсветку
            RestoreHighlightState();

            // 4) Пересчитываем физические суммы
            RecalculatePhysicalSums();
        }


        /// <summary>
        /// Обновляет _savedHighlightState:
        /// - Если в дереве есть узлы, которые были ранее подсвечены и всё ещё подсвечены, оставляет их.
        /// - Если в дереве есть узлы, которые утратили подсветку, удаляет их из словаря.
        /// - Если в дереве появились новые подсвеченные узлы, добавляет их с текущим цветом.
        /// - Узлы, отсутствующие в дереве, в словаре не трогает.
        /// </summary>
        private void SaveHighlightState()
        {
            // 1) Собираем временную карту { имя → цвет } для всех текущих подсвеченных узлов
            var currentState = new Dictionary<string, ColorCircle>();
            foreach (var node in _treeNodes)
                CollectState(node, currentState);

            // 2) Собираем множество имён всех узлов, которые есть сейчас в дереве
            var treeNames = new HashSet<string>();
            void CollectNames(SystemNodeViewModel n)
            {
                treeNames.Add(n.Name);
                foreach (var ch in n.Children)
                    CollectNames(ch);
            }
            foreach (var root in _treeNodes)
                CollectNames(root);

            // 3) Из словаря удаляем те узлы, которые есть в дереве, но уже не подсвечены
            var keys = _savedHighlightState.Keys.ToList();
            foreach (var name in keys)
            {
                if (treeNames.Contains(name) && !currentState.ContainsKey(name))
                    _savedHighlightState.Remove(name);
            }

            // 4) Добавляем новые подсвеченные узлы
            foreach (var kvp in currentState)
            {
                if (!_savedHighlightState.ContainsKey(kvp.Key))
                    _savedHighlightState[kvp.Key] = kvp.Value;
            }
        }

        /// <summary>
        /// Рекурсивно собирает в state пары { node.Name → цвет } для всех узлов с IsHighlighted==true.
        /// </summary>
        private void CollectState(SystemNodeViewModel node, Dictionary<string, ColorCircle> state)
        {
            if (node.IsHighlighted && node.NodeBrush is SolidColorBrush sb)
            {
                var c = sb.Color;
                state[node.Name] = ColorCircle.FromArgb(c.A, c.R, c.G, c.B);
            }
            foreach (var child in node.Children)
                CollectState(child, state);
        }

        private void RestoreHighlightState()
        {
            foreach (var node in _treeNodes)
                RestoreState(node);
        }

        private void RestoreState(SystemNodeViewModel node)
        {
            if (_savedHighlightState.TryGetValue(node.Name, out var savedColor))
            {
                // Применяем именно тот цвет, что был до фильтрации
                node.ApplyHighlight(true, savedColor);
            }

            foreach (var child in node.Children)
                RestoreState(child);
        }

        /// <summary>
        /// Вспомогательный метод, который создаёт узел со всеми его детьми (как в BuildTree)
        /// </summary>
        private SystemNodeViewModel CreateNode(SystemClass sys)
        {
            // Собираем все элементы системы через подсистемы и трубопроводы
            var allSystemIds = sys.Subsystems
                .SelectMany(s => s.PipeLines)
                .SelectMany(p => p.Elements)
                .Select(e => (IModelElementId)e.ColorId)
                .ToArray();

            var sysNode = new SystemNodeViewModel(sys.Name, allSystemIds, this);

            foreach (var subs in sys.Subsystems)
            {
                var subIds = subs.PipeLines
                    .SelectMany(p => p.Elements)
                    .Select(e => (IModelElementId)e.ColorId)
                    .ToArray();

                var subNode = new SystemNodeViewModel(subs.Name, subIds, this);

                foreach (var pipe in subs.PipeLines)
                {
                    var pipeElementIds = pipe.Elements
                        .Select(e => (IModelElementId)e.ColorId)
                        .ToArray();

                    var pipeNode = new SystemNodeViewModel(pipe.Name, pipeElementIds, this);
                    subNode.Children.Add(pipeNode);
                }

                sysNode.Children.Add(subNode);
            }

            return sysNode;
        }


        // Рекурсивно обходит дерево и добавляет имя узла, если он подсвечен
        private void CollectHighlighted(SystemNodeViewModel node, HashSet<string> collector)
        {
            if (node.IsHighlighted)
                collector.Add(node.Name);

            foreach (var child in node.Children)
                CollectHighlighted(child, collector);
        }

        #endregion

        #region Покраска элементов

        /// <summary>
        /// Подсветить набор элементов по их IModelElementId
        /// </summary>
        internal void HighlightByElementIds(IModelElementId[] elementIds, ColorCircle color)
        {
            if (ModelViewer == null || elementIds == null || elementIds.Length == 0)
                return;

            // Преобразуем ColorCircle в SDK Color
            var sdkColor = new Ascon.Pilot.Bim.SDK.Color(color.R, color.G, color.B, 255);

            // Передаём готовые объекты IModelElementId
            ModelViewer.SetColor(elementIds, sdkColor);
        }

        /// <summary>
        /// Сбросить цвет (серый) по списку IModelElementId
        /// </summary>
        internal void ClearByElementIds(IModelElementId[] elementIds)
        {
            if (ModelViewer == null || elementIds == null || elementIds.Length == 0)
                return;

            // Серый цвет (можно вынести в константу)
            var gray = new Ascon.Pilot.Bim.SDK.Color(200, 200, 200, 255);

            ModelViewer.SetColor(elementIds, gray);
        }

        // Вспомогательные методы для глобальной перекраски
        public void GrayAllElements()
        {
            if (ModelViewer == null) return;
            // Берём все элементы модели
            var all = ModelViewer.GetVisibleElements().ToList();
            if (!all.Any()) return;

            // Серый
            var gray = new Ascon.Pilot.Bim.SDK.Color(200, 200, 200, 255);
            // Перекрашиваем
            ModelViewer.SetColor(all, gray);
        }

        public void ReturnVisibilityToAll()
        {
            if (ModelViewer == null) return;
            // Берём все элементы модели
            var all = ModelViewer.GetVisibleElements().ToList();
            if (!all.Any()) return;

            ModelViewer.Show(all);
        }

        public void RestoreOriginalColors()
        {
            if (ModelViewer == null) return;
            // Сброс override-ов на все элементы
            var all = ModelViewer.GetVisibleElements().ToList();
            if (!all.Any()) return;

            ModelViewer.ClearColors(all);
        }


        /// <summary>
        /// Выдаёт случайный цвет, которого ещё нет в _usedColors.
        /// </summary>
        public ColorCircle GetNextUniqueColor()
        {
            ColorCircle c;
            do
            {
                c = ColorCircle.FromRgb(
                    (byte)_rnd.Next(0, 256),
                    (byte)_rnd.Next(0, 256),
                    (byte)_rnd.Next(0, 256));
            } while (_usedColors.Contains(c));
            _usedColors.Add(c);
            return c;
        }

        public ICommand SelectAllTopLevelCommand { get; }
        public ICommand ClearAllCommand { get; }

        // Обработка «выделить все верхнеуровневые»
        private void OnSelectAllTopLevel()
        {
            foreach (var node in TreeNodes)
            {
                node.IsHighlighted = true;
            }
            RecalculatePhysicalSums();
        }

        // Обработка «снять все выделения»
        public void OnClearAll()
        {
            // 1) Убираем все галочки и сбрасываем цвет для каждой ветки дерева
            foreach (var node in TreeNodes)
                ClearNodeRecursively(node);

            // 2) Очищаем сохранённое состояние (теперь нет ни одной подсветки)
            _savedHighlightState.Clear();

            // 3) Сбрасываем все override цвета в самом ModelViewer
            GrayAllElements();

            // 4) Пересчитываем физические суммы
            RecalculatePhysicalSums();
        }

        /// <summary>
        /// Рекурсивно сбрасывает подсветку у узла и у его детей.
        /// </summary>
        private void ClearNodeRecursively(SystemNodeViewModel node)
        {
            // ApplyHighlight сбрасывает IsHighlighted и снимает цвет в ModelViewer через ClearByElementIds
            //var defaultColor = ColorCircle.FromArgb(255, 200, 200, 200);
            //node.ApplyHighlight(false, defaultColor);
            node.IsHighlighted = false;

            foreach (var child in node.Children)
                ClearNodeRecursively(child);
        }
        #endregion

        public void OnModelClosed()
        {
            // При закрытии плагина восстанавливаем исходные цвета
            RestoreOriginalColors();

            ModelViewer = null;
            SidebarTab.IsVisible = false;

            
        }

        // Вспомогательный обход: отдаёт все листья под node (PipeLine-узлы)
        private static IEnumerable<SystemNodeViewModel> EnumerateLeaves(SystemNodeViewModel node)
        {
            if (node.Children == null || node.Children.Count == 0)
            {
                yield return node;
                yield break;
            }

            foreach (var child in node.Children)
                foreach (var leaf in EnumerateLeaves(child))
                    yield return leaf;
        }

        public void RecalculatePhysicalSums()
        {
            // Если нужно ещё считать DiaInch, добавьте свойство TotalDiaInch у VM.
            double w = 0, l = 0, v = 0, di = 0;

            if (TreeNodes != null)
            {
                // Берём только выделенные листья (PipeLine)
                var highlightedPipes = TreeNodes
                    .SelectMany(root => EnumerateLeaves(root))
                    .Where(n => n.IsHighlighted);

                foreach (var pipe in highlightedPipes)
                {
                    w += Convert.ToDouble(pipe.Totals.Weight);
                    l += Convert.ToDouble(pipe.Totals.Length);
                    v += Convert.ToDouble(pipe.Totals.Volume);
                    di += Convert.ToDouble(pipe.Totals.DiaInch);
                }
            }

            // Присваиваем суммарные значения
            TotalTonnage = w;
            TotalLength = l;
            TotalVolume = v;

            //ShowHighlightTree();
        }

        public void ShowHighlightTree()
        {
            if (TreeNodes == null || TreeNodes.Count == 0)
            {
                MessageBox.Show("Дерево пустое.");
                return;
            }

            var sb = new StringBuilder();

            foreach (var root in TreeNodes)
                AppendNodeText(root, sb, 0);

            MessageBox.Show(sb.ToString(), "Подсветка дерева", MessageBoxButton.OK, MessageBoxImage.Information);
        }

        private void AppendNodeText(SystemNodeViewModel node, StringBuilder sb, int indent)
        {
            // Отступ для визуализации уровня вложенности
            sb.Append(new string(' ', indent * 2));

            // Формат: [X] Имя (суммы)
            var flag = node.IsHighlighted ? "[X]" : "[ ]";
            sb.AppendLine($"{flag} {node.Title}");

            // Рекурсия в детей
            foreach (var child in node.Children)
                AppendNodeText(child, sb, indent + 1);
        }


        private DelegateCommand hideUnhighlightedCommand;
        public ICommand HideUnhighlightedCommand
        {
            get
            {
                if (hideUnhighlightedCommand == null)
                {
                    hideUnhighlightedCommand = new DelegateCommand(HideUnhighlighted);
                }

                return hideUnhighlightedCommand;
            }
        }


        private DelegateCommand restoreVisibilityCommand;
        public ICommand RestoreVisibilityCommand
        {
            get
            {
                if (restoreVisibilityCommand == null)
                {
                    restoreVisibilityCommand = new DelegateCommand(ReturnVisibilityForAll);
                }

                return restoreVisibilityCommand;
            }
        }

        public void HideUnhighlighted()
        {
            HideUnhighlighted_mode = true;

            foreach (var node in TreeNodes)
            {
                foreach (var child in node.Children)
                {
                    foreach(var child2 in child.Children)
                    {
                        if (!child2.IsHighlighted)
                        {
                            child2.HideUnhighlighted(ModelViewer);
                        }
                    }
                }
            }
        }

        public void ReturnVisibilityForAll()
        {
            HideUnhighlighted_mode = false;

            foreach (var node in TreeNodes)
            {
                foreach(var child in node.Children)
                {
                    foreach (var child2 in child.Children)
                    {
                        if (!child2.IsHighlighted) child2.ReturnVisibility(ModelViewer);
                    }
                }
                
            }
        }

        public IModelViewer GetModelViewer() {  return ModelViewer; }

        #region NotifyPropertyChanged

        public event PropertyChangedEventHandler PropertyChanged;
        internal virtual void OnPropertyChanged([CallerMemberName] string propertyName = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        #endregion

        #region Мониторинг и контроль испытаний

        #region ФО и методы на создание легенд
       
        private double TP_totalTonnage;
        public double TP_TotalTonnage
        {
            get => TP_totalTonnage;
            set
            {
                if (Math.Abs(TP_totalTonnage - value) < 0.01) return;
                TP_totalTonnage = value;
                OnPropertyChanged();
            }
        }

        private double TP_totalLength;
        public double TP_TotalLength
        {
            get => TP_totalLength;
            set
            {
                if (Math.Abs(TP_totalLength - value) < 0.01) return;
                TP_totalLength = value;
                OnPropertyChanged();
            }
        }

        private double TP_totalWeld;
        public double TP_TotalWeld
        {
            get => TP_totalWeld;
            set
            {
                if (Math.Abs(TP_totalWeld - value) < 0.01) return;
                TP_totalWeld = value;
                OnPropertyChanged();
            }
        }


        public DelegateCommand TP_ShowStatusCommand { get; }
        public DelegateCommand TP_ShowTypeCommand { get; }

        private void TP_ShowStatus()
        {
            if (LoadStatus != "Все данные готовы к визуализации")
            {
                MessageBox.Show("Дождитесь полной загрузки данных!", "", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            TP_Legend.Clear();
            //GrayAllElements();
            OnClearAll();

            var statuses = BuildStatusIndex(TestData);

            var order = statuses.Keys.ToList();
            order.Sort();

            foreach (var rawStatus in order)
            {
                var normalizedStatus = NormalizeStatus(rawStatus);
                var elementIds = statuses[rawStatus].ToArray();

                ColorCircle color;
                string label = normalizedStatus;

                if (normalizedStatus == "TP OFFICE-WELDLOG OK")
                    color = ColorCircle.FromRgb(160, 210, 160);      // яркий пастельный зелёный (положительный)
                else if (normalizedStatus.Contains("TP OFFICE-NOT SUBMITTED"))
                    color = ColorCircle.FromRgb(180, 150, 150);      // (отрицательный)
                else if (normalizedStatus.Contains("TP OFFICE-A PUNCH") || normalizedStatus == "TP-OFFICE A PUNCH")
                    color = ColorCircle.FromRgb(120, 140, 180);      // мягкий синий (отрицательный)
                else if (normalizedStatus == "TP OFFICE")
                    color = ColorCircle.FromRgb(150, 180, 200);      // светлый голубой (нейтральный)

                else if (normalizedStatus.Contains("USTAY QC-PUNCH") || normalizedStatus == "USTAY QC PUNCH")
                    color = ColorCircle.FromRgb(140, 120, 160);      // светлый фиолетовый (отрицательный)
                else if (normalizedStatus.Contains("USTAY QC-WELDLOG"))
                    color = ColorCircle.FromRgb(130, 150, 190);      // пастельный синий (отрицательный)
                else if (normalizedStatus == "USTAY QC-WELDING SUMMARY")
                    color = ColorCircle.FromRgb(170, 190, 220);      // очень светлый голубой (положительный)
                else if (normalizedStatus == "USTAY QC")
                    color = ColorCircle.FromRgb(150, 200, 190);      // яркий бирюзовый (положительный)

                else if (normalizedStatus.Contains("GAZPROM ENG."))
                    color = ColorCircle.FromRgb(220, 210, 160);      // светлый пастельный жёлтый (положительный)

                else if (normalizedStatus == "PRECOM.")
                    color = ColorCircle.FromRgb(140, 210, 200);      // яркий зелёно-бирюзовый (положительный)

                else if (normalizedStatus == "TESTED")
                    color = ColorCircle.FromRgb(120, 170, 220);      // яркий пастельный голубой (положительный)

                else if (normalizedStatus == "SG-1C")
                    color = ColorCircle.FromRgb(195, 170, 205);      // светлый лавандовый (нейтральный)
                else if (normalizedStatus == "BEKIR SEF")
                    color = ColorCircle.FromRgb(180, 190, 170);      // светлый оливковый (нейтральный)
                else if (normalizedStatus == "GOKHAN")
                    color = ColorCircle.FromRgb(200, 180, 170);      // светлый персиковый (нейтральный)
                else if (normalizedStatus == "MATAN")
                    color = ColorCircle.FromRgb(230, 220, 200);      // светлый бежевый (нейтральный)

                else if (normalizedStatus == "RFT")
                    color = ColorCircle.FromRgb(130, 160, 130);      // (нейтральный)

                else
                    color = ColorCircle.FromRgb(180, 180, 180);      // светло-серый (по умолчанию, отрицательный)



                HighlightByElementIds(elementIds, color);

                var wpfColor = System.Windows.Media.Color.FromArgb(255, color.R, color.G, color.B);
                var brush = new System.Windows.Media.SolidColorBrush(wpfColor);

                if (label == "НЕ ОПРЕДЕЛЁН") label = "Не определён";

                var vm = new TP_LegendItem(
                label,
                brush,
                elementIds,
                this,
                _modelViewer,
                false);

                TP_Legend.Add(vm);
            }

            TP_RecalculatePhysicalSums();
        }

        private void TP_ShowType()
        {
            if (LoadStatus != "Все данные готовы к визуализации")
            {
                MessageBox.Show("Дождитесь полной загрузки данных!", "", MessageBoxButton.OK, MessageBoxImage.Exclamation);
                return;
            }

            TP_Legend.Clear(); // Очищаем старые элементы легенды
            OnClearAll();

            var tests = BuildTestTypeIndex(TestData);

            

            var order = tests.Keys.ToList();
            order.Sort();

            foreach (var testType in order)
            {
                var upperType = !string.IsNullOrWhiteSpace(testType)
                    ? testType.ToUpperInvariant()
                    : "Не определён";

                var elementIds = tests[testType].Cast<IModelElementId>().ToArray();

                ColorCircle circleColor;
                string label;

                if (upperType == "HYDRO")
                {
                    circleColor = ColorCircle.FromRgb(0, 0, 255);
                    label = "Гидро";
                }
                else if (upperType == "VISUAL")
                {
                    circleColor = ColorCircle.FromRgb(0, 255, 0);
                    label = "Визуальные";
                }
                else if (upperType == "PNEUMATIC")
                {
                    circleColor = ColorCircle.FromRgb(255, 0, 0);
                    label = "Пневматические";
                }
                else
                {
                    circleColor = ColorCircle.FromRgb(222, 222, 222);
                    label = "Не определён";
                }

                // Окраска в 3D
                HighlightByElementIds(elementIds, circleColor);

                // Добавление в легенду
                var wpfColor = System.Windows.Media.Color.FromArgb(255, circleColor.R, circleColor.G, circleColor.B);
                var brush = new System.Windows.Media.SolidColorBrush(wpfColor);

                var vm = new TP_LegendItem(
                label,
                brush,
                elementIds,
                this,
                _modelViewer,
                false);

                TP_Legend.Add(vm);
            }

            TP_RecalculatePhysicalSums();
        }


        /// <summary>
        /// Группирует все существующие записи TPData по полю TestType.
        /// </summary>
        /// <returns>
        /// Словарь: ключ — TestType (или "Unknown"), 
        /// значение — список ForColorModelElementId соответствующих записей.
        /// </returns>
        public static Dictionary<string, List<ForColorModelElementId>> BuildTestTypeIndex(TPData data)
        {
            return data.Items
                .Values
                .GroupBy(info =>
                    string.IsNullOrWhiteSpace(info.TestType)
                        ? "Не определён"
                        : info.TestType.Trim()
                )
                .ToDictionary(
                    grp => grp.Key,
                    grp => grp.SelectMany(info => info.ModelElementIds).ToList()
                );
        }

        public static Dictionary<string, List<IModelElementId>> BuildStatusIndex(TPData data)
        {
            return data.Items
                    .Values
                    .GroupBy(info =>
                        string.IsNullOrWhiteSpace(info.TPStatus)
                            ? "Не определён"
                            : info.TPStatus.Trim().ToUpperInvariant()
                    )
                    .ToDictionary(
                        grp => grp.Key,
                        grp => grp.SelectMany(info => info.ModelElementIds.Select(id => (IModelElementId)id)).ToList()
                    );
        }

        private string NormalizeStatus(string status)
        {
            if (string.IsNullOrWhiteSpace(status))
                return "Не определён";

            return status.Trim().ToUpperInvariant();
        }


        public void LegendItem_Click(object sender, MouseButtonEventArgs e)
        {
            if (sender is StackPanel panel && panel.DataContext is TP_LegendItem item)
            {
                // Обработка нажатия по item
                MessageBox.Show($"Вы нажали на: {item.Label}");
                if (item.Label == "Системы")
                {
                    OnClearAll();
                    ReturnVisibilityForAll();
                }
                else if (item.Label == "Мониторинг")
                {
                    PerformRestoreVisibility_Monitoring();
                    PerformClearAll_Monitoring();
                }
                else if (item.Label == "Паспорта")
                {
                    PerformRestoreVisibility_Passsports();
                    PerformClearAll_Passports();
                }
                else if (item.Label == "Данные")
                {
                    
                }
            }
        }

        public void TP_RecalculatePhysicalSums()
        {
            double totalLength = 0;
            double totalTonnage = 0;
            double totalWELD = 0;

            var used_guid = new Dictionary<Guid,  int>();

            // Проходим по всем включённым пунктам легенды
            foreach (var item in TP_Legend.Where(x => x.IsVisible))
            {
                foreach (var modelId in item.ElementIds.GroupBy(e => e.ElementId)
                                                                .Select(g => g.First())
                                                                .ToArray())
                {
                    // Пытаемся извлечь Guid из IModelElementId
                    // Предполагаем, что runtime-тип — ForColorModelElementId с публичным Guid-полем/свойством Id
                    if (modelId is ForColorModelElementId fcId)
                    {
                        if (TestData.Items.TryGetValue(fcId.ElementId, out var info) && !used_guid.ContainsKey(fcId.ElementId))
                        {
                            // аккумулируем, учитывая, что в TechnicalInfo поля nullable
                            totalLength += info.TotalTpLengthMeter ?? 0;
                            totalTonnage += info.TotalTpWeightKg / 1000 ?? 0;
                            totalWELD += info.TotalTpWdi ?? 0;
                            //used_guid[fcId.ElementId] = 1;
                            //react_guid += fcId.ToString() + "\n";
                        }
                    }
                }
            }

            //MessageBox.Show(fact_guid + "\n\n" + react_guid);

            TP_TotalLength = totalLength;
            TP_TotalTonnage = totalTonnage;
            TP_TotalWeld = totalWELD;
        }

        public void TP_RecalculateByElementID(IModelElementId[] ElementIds) 
        {
            double totalLength = 0;
            double totalTonnage = 0;
            double totalWELD = 0;


            foreach (var modelId in ElementIds)
            {
                // Пытаемся извлечь Guid из IModelElementId
                // Предполагаем, что runtime-тип — ForColorModelElementId с публичным Guid-полем/свойством Id
                if (modelId is ForColorModelElementId fcId)
                {
                    if (TestData.Items.TryGetValue(fcId.ElementId, out var info))
                    {
                        // аккумулируем, учитывая, что в TechnicalInfo поля nullable
                        totalLength += info.TotalTpLengthMeter ?? 0;
                        totalTonnage += info.TotalTpWeightKg / 1000 ?? 0;
                        totalWELD += info.TotalTpWdi ?? 0;
                    }
                }
            }


            TP_TotalLength = totalLength;
            TP_TotalTonnage = totalTonnage;
            TP_TotalWeld = totalWELD;
        }

        #endregion


        #region Slider
        // --- дата‑поле, выбранное в ComboBox ---
        private string _selectedDateField;
        public string SelectedDateField
        {
            get => _selectedDateField;
            set
            {
                if (_selectedDateField == value) return;
                _selectedDateField = value;
                OnPropertyChanged(nameof(SelectedDateField));
                BuildDateBuckets();               // пересоздаём 10 порогов
                SelectedBucketIndex = 0;          // сбрасываем слайдер на первый
                SelectedDate = _dateBuckets.Count > 0 ? _dateBuckets[SelectedBucketIndex].Threshold : Convert.ToDateTime("01.01.2000");
                HighlightCurrentBucket();
            }
        }

        // Список названий полей DateTime? в TechnicalInfo
        public List<string> DateFieldOptions { get; private set; } = new List<string>();

        private void BuildDateFieldOptions()
        {
            if (TestData is null) return;

            var allFields = new List<string>
            {
                nameof(TechnicalInfo.TestLimitsPreparationDate),
                nameof(TechnicalInfo.SubmittedToGazpromEng),
                nameof(TechnicalInfo.GazpromEngToUstayDate),
                nameof(TechnicalInfo.ToUstayQcForPunch),
                nameof(TechnicalInfo.QcToTpOffice),
                nameof(TechnicalInfo.UstayQcRftAppDate),
                nameof(TechnicalInfo.GazpromQcAppDate),
                nameof(TechnicalInfo.TestDate)
            };

            DateFieldOptions = allFields
                .Where(fieldName =>
                    TestData.Items.Values.Any(info =>
                    {
                        var pi = typeof(TechnicalInfo).GetProperty(fieldName);
                        var val = (DateTime?)pi.GetValue(info);
                        return val.HasValue;
                    }))
                .ToList();

            OnPropertyChanged(nameof(DateFieldOptions));
        }

        // Сами «ведра» порогов
        private List<DateThresholdBucket> _dateBuckets = new List<DateThresholdBucket>();
        public IReadOnlyList<DateThresholdBucket> DateBuckets => _dateBuckets;

        // индекс выбранного «ведра»
        private int _selectedBucketIndex;
        public int SelectedBucketIndex
        {
            get => _selectedBucketIndex;
            set
            {
                if (_selectedBucketIndex == value || _dateBuckets.Count == 0) return;
                _selectedBucketIndex = value;
                OnPropertyChanged(nameof(SelectedBucketIndex));

                // при смене индекса — показываем дату и перекрашиваем
                SelectedDate = _dateBuckets[_selectedBucketIndex].Threshold;
                HighlightCurrentBucket();
            }
        }

        // для отображения в TextBlock
        private DateTime _selectedDate;
        public DateTime SelectedDate
        {
            get => _selectedDate;
            private set
            {
                if (_selectedDate == value) return;
                _selectedDate = value;
                OnPropertyChanged(nameof(SelectedDate));
            }
        }

        // диапозон слайдера 0 … Count-1
        public double MinDateTicks => 0;
        public double MaxDateTicks => Math.Max(0, _dateBuckets.Count - 1);
        public double TickStep => 1;

        // … остальной код VM: TP_Legend, TestData, HighlightByElementIds/ClearByElementIds …

        // Вызывается из сеттера SelectedDateField
        private void BuildDateBuckets()
        {
            if (TestData is null) return;

            _dateBuckets.Clear();

            // 1) собираем все непустые даты из TestData по выбранному полю
            var dates = TestData.Items.Values
                .Select(info =>
                {
                    var pi = typeof(TechnicalInfo).GetProperty(SelectedDateField);
                    return (DateTime?)pi.GetValue(info);
                })
                .Where(d => d.HasValue)
                .Select(d => d.Value)
                .OrderBy(d => d)
                .ToList();

            if (!dates.Any())
            {
                OnPropertyChanged(nameof(MaxDateTicks));
                return;
            }

            var min = dates.First();
            var max = dates.Last();
            var span = max - min;

            // 2) создаём 10 пороговых дат (равномерно от min до max)
            for (int i = 0; i < 10; i++)
            {
                var threshold = min + TimeSpan.FromTicks(span.Ticks * i / 9);
                // 3) для каждого порога собираем все IModelElementId,
                //    у которых дата <= threshold
                var ids = TestData.Items.Values
                    .Where(info =>
                    {
                        var val = (DateTime?)typeof(TechnicalInfo)
                                    .GetProperty(SelectedDateField)
                                    .GetValue(info);
                        return val.HasValue && val.Value <= threshold;
                    })
                    .SelectMany(info => info.ModelElementIds.Select(id => (IModelElementId)id))
                    .ToArray();
                
                _dateBuckets.Add(new DateThresholdBucket(threshold, ids));
            }

            // чтобы обновить слайдер в UI
            OnPropertyChanged(nameof(MaxDateTicks));
        }

        // подсвечиваем текущий набор элементом зелёным
        private void HighlightCurrentBucket()
        {
            // 1) сначала очищаем предыдущую окраску по всем
            GrayAllElements(); // метод, снимающий подсветку у всех
                               // 2) красим в зелёный

            if (_dateBuckets.Count == 0) return;
            var bucket = _dateBuckets[_selectedBucketIndex];
            foreach (var id in bucket.ElementIds)
            {
                HighlightByElementIds(new[] { id },
                    ColorCircle.FromArgb(255, 0, 200, 0)); // зелёный
            }
            TP_RecalculateByElementID(bucket.ElementIds);
        }
        #endregion

        #region Кнопки для быстрого выделения

        private DelegateCommand selectAll_Monitoring;
        public ICommand SelectAll_Monitoring
        {
            get
            {
                if (selectAll_Monitoring == null)
                {
                    selectAll_Monitoring = new DelegateCommand(PerformSelectAll_Monitoring);
                }

                return selectAll_Monitoring;
            }
        }

        private void PerformSelectAll_Monitoring()
        {
            foreach(var legend_item in TP_Legend)
            {
                legend_item.IsVisible = true;
            }
        }

        private DelegateCommand clearAll_Monitoring;
        public ICommand ClearAll_Monitoring
        {
            get
            {
                if (clearAll_Monitoring == null)
                {
                    clearAll_Monitoring = new DelegateCommand(PerformClearAll_Monitoring);
                }

                return clearAll_Monitoring;
            }
        }

        public void PerformClearAll_Monitoring()
        {
            foreach(var legend_item in TP_Legend)
            {
                legend_item.IsVisible = false;
            }
        }

        private DelegateCommand hideUnhighlighted_Monitoring;
        public ICommand HideUnhighlighted_Monitoring
        {
            get
            {
                if (hideUnhighlighted_Monitoring == null)
                {
                    hideUnhighlighted_Monitoring = new DelegateCommand(PerformHideUnhighlighted_Monitoring);
                }

                return hideUnhighlighted_Monitoring;
            }
        }

        private void PerformHideUnhighlighted_Monitoring()
        {
            foreach(var legend_elem in TP_Legend)
            {
                legend_elem.HideU(true);
                if (legend_elem.IsVisible == false)
                {
                    ModelViewer.Hide(legend_elem.ElementIds.GroupBy(e => e.ElementId)
                                                            .Select(g => g.First())
                                                            .ToArray());
                }
                else continue;
            }
        }

        private DelegateCommand restoreVisibility_Monitoring;
        public ICommand RestoreVisibility_Monitoring
        {
            get
            {
                if (restoreVisibility_Monitoring == null)
                {
                    restoreVisibility_Monitoring = new DelegateCommand(PerformRestoreVisibility_Monitoring);
                }

                return restoreVisibility_Monitoring;
            }
        }

        public void PerformRestoreVisibility_Monitoring()
        {
            foreach (var legend_elem in TP_Legend)
            {
                legend_elem.HideU(false);
                if (legend_elem.IsVisible == false)
                {
                    ModelViewer.Show(legend_elem.ElementIds.GroupBy(e => e.ElementId)
                                                            .Select(g => g.First())
                                                            .ToArray());
                }
                else continue;
            }
        }

        #endregion

        #endregion

        #region Управление ручной загрузкой данных
        public void get_filepath(string filePath)
        {
            _filePath = filePath;
        }

        private DelegateCommand loadModelPart;
        public ICommand LoadModelPart
        {
            get
            {
                if (loadModelPart == null)
                {
                    loadModelPart = new DelegateCommand(PerformLoadModelPart);
                }

                return loadModelPart;
            }
        }

        private async void PerformLoadModelPart()
        {
            if (LoadedModelParts.Count == 0)
            {
                MessageBox.Show("Добавьте хотя бы одну часть модели!", "", MessageBoxButton.OK, MessageBoxImage.Stop);
                //var visiblePartGuids = ModelViewer
                //.GetVisibleElements()
                //.Select(elem => elem.ModelPartId)
                //.Distinct()
                //.ToList();

                return;
            }

            IsLoading = true;

            


             LoadStatus = "Пожалуйста, подождите. Идёт загрузка систем и подсистем...";
            await InitModelDataAsync();

            ApplyFilter(null);

            LoadStatus = "Пожалуйста, подождите. Идёт загрузка данных о статусах подсистем...";
            await InitPhysDataAsync();

            LoadStatus = "Пожалуйста, подождите. Идёт загрузка данных о испытаниях...";
            await InitTPDataAsync();

            LoadStatus = "Все данные готовы к визуализации";
            IsLoading = false;

            //TestData.ShowTPDataHeader();
        }


        private string _selectedModelPartsText;
        public string SelectedModelPartsText
        {
            get => _selectedModelPartsText;
            set
            {
                _selectedModelPartsText = value;
                OnPropertyChanged();
            }
        }


        private DelegateCommand addModelPartCommand;
        public ICommand AddModelPartCommand
        {
            get
            {
                if (addModelPartCommand == null)
                {
                    addModelPartCommand = new DelegateCommand(AddModelPart);
                }

                return addModelPartCommand;
            }
        }

        private void AddModelPart()
        {
            if (cmbModelPartsToLoadSelected == null || cmbModelPartsToLoadSelected == "")
            {
                // PARTMAP
                var partMap = partMap_guid_string;
                //var partMap = new Dictionary<Guid, string>
                //{
                //    { Guid.Parse("1d65306d-d5e3-492e-bb13-9e3327e0d67c"), "900" },
                //    { Guid.Parse("eaad9dd1-96ad-4ba8-9b45-d3abae092a89"), "200" },
                //    { Guid.Parse("138078c3-7b42-4555-91bc-d6c1211107b4"), "010" },
                //    { Guid.Parse("389deea2-ed77-413f-b110-57859325af38"), "300" },
                //    { Guid.Parse("3fb1ab10-3a28-4768-b7b5-a134053217f6"), "310" },
                //    { Guid.Parse("6c6abc95-a485-4aa0-8976-77ea17aac71b"), "320" },
                //    { Guid.Parse("254abf4e-acb5-461d-9339-0b3110e3eb60"), "330" },
                //    { Guid.Parse("29eb450d-4722-41bd-a60a-8dd8ddf0edf2"), "340" },
                //    { Guid.Parse("92e14574-3cf5-4bc0-af26-1c7762295b61"), "410" },
                //    { Guid.Parse("57f2df13-e310-4584-9adc-1babbce5057c"), "420" },
                //    { Guid.Parse("62b19655-c26b-4b28-92b7-c5568844e845"), "500" },
                //    { Guid.Parse("49107e4d-7f91-4dcc-9230-082d6c4af67a"), "720" },
                //    { Guid.Parse("971e43f1-a839-4e22-91c3-78240ed48e40"), "730" },
                //    { Guid.Parse("29c44c09-e9c1-414f-b3ae-07dbb271b53b"), "801" },
                //    { Guid.Parse("caf32630-af08-4ca9-8cc7-76807df05aa9"), "802" },
                //    { Guid.Parse("253af425-1622-4771-85d0-5f74fd9bf06f"), "803" },
                //    { Guid.Parse("7b256da3-fb91-4be6-8836-26b01d709b33"), "804" }
                //};


                foreach (var guid in partMap.Keys)
                {
                    if (_modelViewer.IsModelPartLoaded(guid))
                    {
                        LoadedModelParts.Add(partMap[guid]);
                    }
                }
            }

            //MessageBox.Show("Начал загрузку части: " + cmbModelPartsToLoadSelected);

            if (LoadedModelParts.Contains(cmbModelPartsToLoadSelected)) return;
            LoadedModelParts.Add(cmbModelPartsToLoadSelected);
            SelectedModelPartsText = string.Join(Environment.NewLine, LoadedModelParts);
        }
        #endregion


        #region Создание отчётов
        private DelegateCommand createReportSystems;
        public ICommand CreateReportSystems
        {
            get
            {
                if (createReportSystems == null)
                {
                    createReportSystems = new DelegateCommand(async () => await PerformCreateReportSystemsAsync());
                }

                return createReportSystems;
            }
        }
        /*
        private async Task PerformCreateReportSystemsAsync()
        {
            //await PerformCreateReportSystemsAsync();
            // Локальный обход: перечисляет выделенные ЛИСТЬЯ (PipeLine) с путём System/SubSystem и их Totals
            IEnumerable<(string System, string Subsystem, string Line, PhysicalTotals Totals)> EnumHighlightedLeaves()
            {
                if (TreeNodes == null) yield break;

                foreach (var sysNode in TreeNodes)
                {
                    // system = имя корневого узла
                    foreach (var item in Recurse(sysNode, system: sysNode.Name, subsystem: null))
                        yield return item;
                }

                IEnumerable<(string System, string Subsystem, string Line, PhysicalTotals Totals)>
                    Recurse(SystemNodeViewModel node, string system, string subsystem)
                {
                    // Лист (PipeLine)
                    if (node.Children == null || node.Children.Count == 0)
                    {
                        if (node.IsHighlighted)
                            yield return (system, subsystem ?? string.Empty, node.Name, node.Totals);
                        yield break;
                    }

                    // Первый уровень под системой — это подсистема (фиксируем её имя)
                    if (subsystem == null)
                    {
                        foreach (var child in node.Children)
                            foreach (var t in Recurse(child, system, subsystem: child.Name))
                                yield return t;
                    }
                    else
                    {
                        // Более глубокие уровни: подсистема не меняется
                        foreach (var child in node.Children)
                            foreach (var t in Recurse(child, system, subsystem))
                                yield return t;
                    }
                }
            }

            // Собираем строки отчёта прямо из выделенных листьев
            var rows = EnumHighlightedLeaves()
                .Select(x => new SvodkaRow
                {
                    System = x.System,
                    Subsystem = x.Subsystem,
                    Line = x.Line,
                    // Берём ФО из узла:
                    Length = Convert.ToDouble(x.Totals.Length),
                    Tonnage = Convert.ToDouble(x.Totals.Weight),
                    Volume = Convert.ToDouble(x.Totals.Volume)
                })
                .ToList();

            if (rows.Count == 0)
            {
                if (!AskContinueWhenEmpty())
                    return;
            }

            // Диалог сохранения
            var sfd = new SaveFileDialog
            {
                Title = "Сохранить отчёт 'Сводка'",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx",
                FileName = $"Сводка_{DateTime.Now:yyyyMMdd_HHmm}.xlsx"
            };
            if (sfd.ShowDialog() != true)
                return;

            var wb = new XLWorkbook();
            var ws = wb.Worksheets.Add("Сводка");

            // Заголовки
            var headers = new[] { "Система", "Подсистема", "Линия", "Протяженность", "Тоннаж", "Объём" };
            for (int i = 0; i < headers.Length; i++)
                ws.Cell(1, i + 1).Value = headers[i];

            // Данные
            int row = 2;
            foreach (var r in rows)
            {
                ws.Cell(row, 1).Value = r.System;
                ws.Cell(row, 2).Value = r.Subsystem;
                ws.Cell(row, 3).Value = r.Line;
                ws.Cell(row, 4).Value = r.Length;
                ws.Cell(row, 5).Value = r.Tonnage;
                ws.Cell(row, 6).Value = r.Volume;
                row++;
            }

            // Таблица + итоги
            int lastDataRow = Math.Max(row - 1, 1);
            var rng = ws.Range(1, 1, lastDataRow, 6);
            var table = rng.CreateTable();
            table.Theme = XLTableTheme.TableStyleLight9;

            table.ShowTotalsRow = true;
            table.Field(0).TotalsRowLabel = "ИТОГО:";
            table.Field(3).TotalsRowFunction = XLTotalsRowFunction.Sum; // Протяженность
            table.Field(4).TotalsRowFunction = XLTotalsRowFunction.Sum; // Тоннаж
            table.Field(5).TotalsRowFunction = XLTotalsRowFunction.Sum; // Объём

            // Форматы
            ws.Column(4).Style.NumberFormat.Format = "#,##0.00";
            ws.Column(5).Style.NumberFormat.Format = "#,##0.00";
            ws.Column(6).Style.NumberFormat.Format = "#,##0.00";

            ws.SheetView.FreezeRows(1);
            ws.Columns(1, 6).AdjustToContents();

            wb.SaveAs(sfd.FileName);
        }
        */
        private async Task PerformCreateReportSystemsAsync()
        {
            double totalLength = 0, totalWeight = 0, totalVolume = 0;

            // 1) Сбор исходных данных из дерева — на UI-потоке (без блокировки UI окна прогресса).
            var leaves = new List<(string System, string Subsystem, string Line, PhysicalTotals Totals)>();
            foreach (var x in EnumHighlightedLeaves()) // ваша локальная функция остаётся как есть
                leaves.Add(x);

            if (leaves.Count == 0)
            {
                if (!AskContinueWhenEmpty())
                    return;
            }

            // 2) Диалог выбора файла — до запуска фоновой задачи
            var sfd = new SaveFileDialog
            {
                Title = "Сохранить отчёт 'Сводка'",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx|PDF Document (*.pdf)|*.pdf",
                FileName = $"Сводка_{DateTime.Now:yyyyMMdd_HHmm}"
            };
            if (sfd.ShowDialog() != true)
                return;

            // 3) Окно прогресса
            var dlg = new ProgressDialog { Owner = Application.Current.MainWindow };
            dlg.SetIndeterminate("Подготовка данных...");
            dlg.Show();

            // Прогресс: маршалится в UI-поток автоматически (создан на UI-потоке)
            // Важно: интерфейсный тип слева
            IProgress<ReportProgress> progress =
                new System.Progress<ReportProgress>(p =>
                {
                    // обновление UI
                    dlg.SetProgress(p.Phase, p.Current, p.Total);
                });

            try
            {
                await Task.Run(() =>
                {
                    var token = dlg.Token;
                    token.ThrowIfCancellationRequested();

                    // --- Стадия 1: сбор строк
                    progress.Report(new ReportProgress { Phase = "Сбор данных...", Current = 0, Total = leaves.Count });

                    var sb = new StringBuilder();

                    // Заголовки
                    sb.AppendLine("Система;Подсистема;Линия;Протяженность;Тоннаж;Объём");

                    int i = 0;
                    foreach (var x in leaves)
                    {
                        token.ThrowIfCancellationRequested();

                        var len = Convert.ToDouble(x.Totals.Length);
                        var wgt = Convert.ToDouble(x.Totals.Weight);
                        var vol = Convert.ToDouble(x.Totals.Volume);

                        var row = string.Join(";", new[]
                        {
                            x.System,
                            x.Subsystem ?? string.Empty,
                            x.Line,
                            len.ToString("0.0000"),
                            wgt.ToString("0.0000"),
                            vol.ToString("0.0000")
                        });

                        sb.AppendLine(row);

                        // аккумулируем
                        totalLength += len;
                        totalWeight += wgt;
                        totalVolume += vol;

                        i++;
                        if ((i & 31) == 0 || i == leaves.Count)
                            progress.Report(new ReportProgress { Phase = "Сбор данных...", Current = i, Total = leaves.Count });
                    }

                    // --- Стадия 2–3: формирование Excel напрямую (без CSV)
                    progress.Report(new ReportProgress { Phase = "Формирование файла...", Current = 0, Total = 0 });

                    var wb = new XLWorkbook();
                    var ws = wb.Worksheets.Add("Сводка");

                    // Заголовки
                    ws.Cell(1, 1).Value = "Система";
                    ws.Cell(1, 2).Value = "Подсистема";
                    ws.Cell(1, 3).Value = "Линия";
                    ws.Cell(1, 4).Value = "Протяженность"; // единицы оставляем в PDF-хедере
                    ws.Cell(1, 5).Value = "Тоннаж";
                    ws.Cell(1, 6).Value = "Объём";

                    // Данные (ВАЖНО: числа кладём как double, без ToString)
                    int rowIdx = 2;
                    foreach (var x in leaves)
                    {
                        token.ThrowIfCancellationRequested();

                        ws.Cell(rowIdx, 1).Value = x.System;
                        ws.Cell(rowIdx, 2).Value = x.Subsystem ?? string.Empty;
                        ws.Cell(rowIdx, 3).Value = x.Line;

                        ws.Cell(rowIdx, 4).Value = Convert.ToDouble(x.Totals.Length);
                        ws.Cell(rowIdx, 5).Value = Convert.ToDouble(x.Totals.Weight);
                        ws.Cell(rowIdx, 6).Value = Convert.ToDouble(x.Totals.Volume);

                        rowIdx++;
                    }

                    // Оформление таблицы
                    progress.Report(new ReportProgress { Phase = "Оформление...", Current = 0, Total = 0 });

                    int lastDataRow = rowIdx - 1;
                    var rng = ws.Range(1, 1, lastDataRow, 6);
                    var table = rng.CreateTable();
                    table.Theme = XLTableTheme.TableStyleLight9;

                    // Числовые колонки: явно задаём тип Number и формат
                    var data = table.DataRange; // без заголовка


                    ws.Column(4).Style.NumberFormat.Format = "#,##0.00";
                    ws.Column(5).Style.NumberFormat.Format = "#,##0.00";
                    ws.Column(6).Style.NumberFormat.Format = "#,##0.00";

                    // Итого
                    table.ShowTotalsRow = true;
                    table.Field(0).TotalsRowLabel = "ИТОГО:";
                    table.Field(3).TotalsRowLabel = totalLength.ToString("N2");
                    table.Field(4).TotalsRowLabel = totalWeight.ToString("N2");
                    table.Field(5).TotalsRowLabel = totalVolume.ToString("N2");

                    ws.SheetView.FreezeRows(1);
                    ws.Columns(1, 6).AdjustToContents();


                    // --- Стадия 4: сохранение
                    progress.Report(new ReportProgress { Phase = "Сохранение файла...", Current = 0, Total = 0 });

                    string ext = System.IO.Path.GetExtension(sfd.FileName).ToLowerInvariant();
                    if (ext == ".xlsx")
                    {
                        wb.SaveAs(sfd.FileName);
                    }
                    else if (ext == ".pdf")
                    {
                        // Соберём данные для PDF из листа ws (или из ваших исходных данных — как удобнее)
                        var headers = new[]
                                        {
                                            "Система",
                                            "Подсистема",
                                            "Линия",
                                            "Протяженность, м",
                                            "Тоннаж, т",
                                            "Объём, м³"
                                        };

                        // Считываем таблицу из используемого диапазона (ровно 6 колонок)
                        var used = ws.RangeUsed();
                        var rows = new List<string[]>();

                        int rowIndex = 0;
                        foreach (var row in used.Rows())
                        {
                            // пропускаем первую строку листа (в ней заголовки Excel)
                            if (rowIndex++ == 0)
                                continue;

                            var cells = row.Cells(1, 6)
                                .Select(c => c.GetFormattedString())
                                .ToArray();

                            rows.Add(cells);
                        }

                        string fmt = "#,##0.00";
                        var totalRow = new[]
                        {
                            "ИТОГО:", "", "",
                            totalLength.ToString(fmt),
                            totalWeight.ToString(fmt),
                            totalVolume.ToString(fmt)
                        };

                        if (rows.Count > 0)
                        {
                            int lastIndex = rows.Count - 1;
                            var last = rows[lastIndex];

                            bool lastIsTotal = last.Length > 0 &&
                                string.Equals((last[0] ?? "").Trim(), "ИТОГО:", StringComparison.OrdinalIgnoreCase);

                            if (lastIsTotal)
                            {
                                // если массив короче 6 элементов — расширим и запишем обратно
                                if (last.Length < 6)
                                {
                                    Array.Resize(ref last, 6);
                                    rows[lastIndex] = last; // важно присвоить обратно
                                }

                                last[0] = "ИТОГО:"; last[1] = ""; last[2] = "";
                                last[3] = totalLength.ToString(fmt);
                                last[4] = totalWeight.ToString(fmt);
                                last[5] = totalVolume.ToString(fmt);
                            }
                            else
                            {
                                rows.Add(totalRow);
                            }
                        }
                        else
                        {
                            rows.Add(totalRow);
                        }


                        // Генерация PDF
                        GeneratePdfReport(
                            filePath: sfd.FileName,
                            title: "Отчёт по системам и подсистемам" + " " + ProjectName,
                            createdAt: DateTime.Now,
                            headers: headers,
                            rows: rows
                        );
                    }

                }, dlg.Token);


                dlg.Close();
                dlg.Close();
                MessageBox.Show("Отчёт успешно сформирован.", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);

                // Если PDF — открываем его сразу
                if (System.IO.Path.GetExtension(sfd.FileName).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = sfd.FileName,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Не удалось открыть файл автоматически:\n{ex.Message}", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                dlg.Close();
                MessageBox.Show("Формирование отчёта отменено.", "Отмена", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            catch (Exception ex)
            {
                dlg.Close();
                MessageBox.Show($"Ошибка при формировании отчёта:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }

            IEnumerable<(string System, string Subsystem, string Line, PhysicalTotals Totals)> EnumHighlightedLeaves()
            {
                if (TreeNodes == null) yield break;
                foreach (var sysNode in TreeNodes)
                {
                    foreach (var item in Recurse(sysNode, system: sysNode.Name, subsystem: null))
                        yield return item;
                }

                IEnumerable<(string System, string Subsystem, string Line, PhysicalTotals Totals)>
                    Recurse(SystemNodeViewModel node, string system, string subsystem)
                {
                    if (node.Children == null || node.Children.Count == 0)
                    {
                        if (node.IsHighlighted)
                            yield return (system, subsystem ?? string.Empty, node.Name, node.Totals);
                        yield break;
                    }

                    if (subsystem == null)
                    {
                        foreach (var child in node.Children)
                            foreach (var t in Recurse(child, system, subsystem: child.Name))
                                yield return t;
                    }
                    else
                    {
                        foreach (var child in node.Children)
                            foreach (var t in Recurse(child, system, subsystem))
                                yield return t;
                    }
                }
            }


        }

        private static bool AskContinueWhenEmpty()
        {
            // Если хотите — покажите MessageBox. Возвращаем true, чтобы продолжить и создать пустой файл.
            MessageBox.Show("В дереве ничего не выделено!", "Предупреждение", MessageBoxButton.OK, MessageBoxImage.Stop);
            return true;
        }

        public class SvodkaRow
        {
            public string System { get; set; }
            public string Subsystem { get; set; }
            public string Line { get; set; }
            public double Length { get; set; }
            public double Tonnage { get; set; }
            public double Volume { get; set; }
        }

        public sealed class ReportProgress
        {
            public string Phase { get; set; } = "";
            public int Current { get; set; }
            public int Total { get; set; }
        }

        private DelegateCommand createReportMonitoring;
        public ICommand CreateReportMonitoring
        {
            get
            {
                if (createReportMonitoring == null)
                {
                    createReportMonitoring = new DelegateCommand(async () => await PerformCreateReportMonitoringAsync());
                }

                return createReportMonitoring;
            }
        }

        private async Task PerformCreateReportMonitoringAsync()
        {
            // Аккумуляторы итогов (для PDF и контроля)
            double totalTonnageT = 0.0;     // т
            double totalLengthM = 0.0;     // м
            double totalWeldInch = 0.0;     // inch

            // 1) Проверка данных
            if (TestData == null || TestData.Items == null || TestData.Items.Count == 0)
            {
                MessageBox.Show("Нет данных для отчёта по Мониторингу и испытаниям.", "Пусто",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // 2) Диалог сохранения
            var sfd = new SaveFileDialog
            {
                Title = "Сохранить отчёт 'Мониторинг и испытания'",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx|PDF Document (*.pdf)|*.pdf",
                FileName = $"Мониторинг_{DateTime.Now:yyyyMMdd_HHmm}"
            };
            if (sfd.ShowDialog() != true)
                return;

            // 3) Окно прогресса
            var dlg = new ProgressDialog { Owner = Application.Current.MainWindow };
            dlg.SetIndeterminate("Подготовка данных...");
            dlg.Show();

            // Прогресс
            IProgress<ReportProgress> progress =
                new System.Progress<ReportProgress>(p => dlg.SetProgress(p.Phase, p.Current, p.Total));

            try
            {
                await Task.Run(() =>
                {
                    // Собираем список записей и фильтр по текущей подсветке легенды
                    var allItems = TestData.Items.Values.ToList();

                    // Построим множество разрешённых ElementId из видимых пунктов легенды
                    HashSet<Guid> allowed = null;
                    if (TP_Legend != null && TP_Legend.Any())
                    {
                        var visibleIds = TP_Legend
                            .Where(x => x.IsVisible)                   // учитываем только включённые элементы легенды
                            .SelectMany(x => x.ElementIds ?? Array.Empty<IModelElementId>())
                            .OfType<ForColorModelElementId>()
                            .Select(x => x.ElementId)
                            .ToList();

                        var allIds = TP_Legend
                            .SelectMany(x => x.ElementIds ?? Array.Empty<IModelElementId>())
                            .OfType<ForColorModelElementId>()
                            .Select(x => x.ElementId)
                            .ToList();

                        // Включаем фильтр ТОЛЬКО если подсветка сейчас не «всё»
                        if (visibleIds.Count > 0 && visibleIds.Count < allIds.Count)
                            allowed = new HashSet<Guid>(visibleIds);
                    }

                    // Оставляем только те TP-записи, у которых есть пересечение с разрешёнными ElementId
                    IEnumerable<TechnicalInfo> filteredItems = allItems;
                    if (allowed != null)
                    {
                        filteredItems = allItems.Where(info =>
                            info?.ModelElementIds != null &&
                            info.ModelElementIds.OfType<ForColorModelElementId>()
                                .Any(id => allowed.Contains(id.ElementId)));
                    }

                    var filteredList = filteredItems.ToList();

                    // Группировка по линии
                    var groupsByLine = filteredList
                        .GroupBy(info => NormalizeLine(ResolveLineName(info)))
                        .OrderBy(g => g.Key, StringComparer.OrdinalIgnoreCase)
                        .ToList();


                    var token = dlg.Token;
                    token.ThrowIfCancellationRequested();

                    // --- Стадия 1: сбор строк из TestData
                    progress.Report(new ReportProgress { Phase = "Сбор данных...", Current = 0, Total = TestData.Items.Count });

                    // Подготовим последовательность записей (без LINQ к UI — всё в фоне)
                    var items = TestData.Items.Values.ToList();

                    // --- Стадия 2–3: формирование Excel напрямую (никаких CSV)
                    progress.Report(new ReportProgress { Phase = "Формирование файла...", Current = 0, Total = 0 });

                    var wb = new XLWorkbook();
                    var ws = wb.Worksheets.Add("Мониторинг");

                    // Заголовки (можно с единицами сразу)
                    ws.Cell(1, 1).Value = "Статус испытания";
                    ws.Cell(1, 2).Value = "Тип испытания";
                    ws.Cell(1, 3).Value = "Линия";
                    ws.Cell(1, 4).Value = "Тоннаж, т";
                    ws.Cell(1, 5).Value = "Протяженность, м";
                    ws.Cell(1, 6).Value = "Сварной шов, inch";

                    int row = 2;
                    int processed = 0;

                    progress.Report(new ReportProgress { Phase = "Формирование файла...", Current = 0, Total = groupsByLine.Count });


                    foreach (var g in groupsByLine)
                    {
                        token.ThrowIfCancellationRequested();

                        string line = g.Key;
                        string status = AggregateLabel(g.Select(x => SafeStatus(x?.TPStatus)));
                        string type = AggregateLabel(g.Select(x => string.IsNullOrWhiteSpace(x?.TestType) ? "Не определён" : x.TestType.Trim()));

                        double tonnageT = g.Sum(x => (x?.TotalTpWeightKg ?? 0) / 1000.0); // кг -> т
                        double lengthM = g.Sum(x => (x?.TotalTpLengthMeter ?? 0));
                        double weldInch = g.Sum(x => (x?.TotalTpWdi ?? 0));

                        ws.Cell(row, 1).Value = status;
                        ws.Cell(row, 2).Value = type;
                        ws.Cell(row, 3).Value = line;
                        ws.Cell(row, 4).Value = tonnageT;
                        ws.Cell(row, 5).Value = lengthM;
                        ws.Cell(row, 6).Value = weldInch;

                        totalTonnageT += tonnageT;
                        totalLengthM += lengthM;
                        totalWeldInch += weldInch;

                        row++;
                        processed++;
                        if ((processed & 63) == 0)
                            progress.Report(new ReportProgress { Phase = "Сбор данных...", Current = processed, Total = groupsByLine.Count });
                    }

                    // Оформление таблицы
                    progress.Report(new ReportProgress { Phase = "Оформление...", Current = 0, Total = 0 });

                    int lastDataRow = row - 1;
                    var rng = ws.Range(1, 1, lastDataRow, 6);
                    var table = rng.CreateTable();
                    table.Theme = XLTableTheme.TableStyleLight9;

                    // Числовые колонки: задаём формат
                    ws.Column(4).Style.NumberFormat.Format = "#,##0.000";
                    ws.Column(5).Style.NumberFormat.Format = "#,##0.00";
                    ws.Column(6).Style.NumberFormat.Format = "#,##0.00";

                    // TotalsRow
                    table.ShowTotalsRow = true;
                    table.Field(0).TotalsRowLabel = "ИТОГО:";
                    table.Field(3).TotalsRowLabel = totalTonnageT.ToString("N2"); // Тоннаж
                    table.Field(4).TotalsRowLabel = totalLengthM.ToString("N2"); // Протяженность
                    table.Field(5).TotalsRowLabel = totalWeldInch.ToString("N2"); // Шов

                    ws.SheetView.FreezeRows(1);
                    ws.Columns(1, 6).AdjustToContents();

                    // --- Стадия 4: сохранение
                    progress.Report(new ReportProgress { Phase = "Сохранение файла...", Current = 0, Total = 0 });

                    string ext = System.IO.Path.GetExtension(sfd.FileName).ToLowerInvariant();

                    if (ext == ".xlsx")
                    {
                        wb.SaveAs(sfd.FileName);
                    }
                    else if (ext == ".pdf")
                    {
                        // Готовим данные для PDF напрямую (без чтения из Excel)
                        var headers = new[]
                        {
                            "Статус испытания",
                            "Тип испытания",
                            "Линия",
                            "Тоннаж, т",
                            "Протяженность, м",
                            "Сварной шов, inch"
                        };

                        var rows = new List<string[]>(capacity: groupsByLine.Count + 2);

                        // те же агрегированные данные, что в Excel
                        foreach (var g in groupsByLine)
                        {
                            string line = g.Key;
                            string status = AggregateLabel(g.Select(x => SafeStatus(x?.TPStatus)));
                            string type = AggregateLabel(g.Select(x => string.IsNullOrWhiteSpace(x?.TestType) ? "Не определён" : x.TestType.Trim()));

                            double tonnageT = g.Sum(x => (x?.TotalTpWeightKg ?? 0) / 1000.0);
                            double lengthM = g.Sum(x => (x?.TotalTpLengthMeter ?? 0));
                            double weldInch = g.Sum(x => (x?.TotalTpWdi ?? 0));

                            rows.Add(new[]
                            {
                                    status,
                                    type,
                                    line,
                                    tonnageT.ToString("#,##0.000"),
                                    lengthM.ToString("#,##0.00"),
                                    weldInch.ToString("#,##0.00")
                                });
                        }

                        // Добавляем «ИТОГО»
                        rows.Add(new[]
                        {
                            "ИТОГО:", "", "",
                            totalTonnageT.ToString("#,##0.000"),
                            totalLengthM.ToString("#,##0.00"),
                            totalWeldInch.ToString("#,##0.00")
                        });


                        // Рендер PDF
                        GeneratePdfReport(
                            filePath: sfd.FileName,
                            title: "Отчёт по мониторингу и контролю испытаний" + " " + ProjectName,
                            createdAt: DateTime.Now,
                            headers: headers,
                            rows: rows
                        );
                    }

                }, dlg.Token);

                dlg.Close();
                MessageBox.Show("Отчёт успешно сформирован.", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);

                // Авто-открытие PDF (как раньше)
                if (System.IO.Path.GetExtension(sfd.FileName).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = sfd.FileName,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Не удалось открыть PDF автоматически:\n{ex.Message}", "Предупреждение",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                dlg.Close();
                MessageBox.Show("Формирование отчёта отменено.", "Отмена", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            catch (Exception ex)
            {
                dlg.Close();
                MessageBox.Show($"Ошибка при формировании отчёта:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        // --- помощники ---

        // Локальные помощники для агрегирования подписей
        static string AggregateLabel(IEnumerable<string> vals)
        {
            var distinct = vals
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .Select(s => s.Trim())
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            if (distinct.Count == 0) return "—";
            if (distinct.Count == 1) return distinct[0];
            // Если статусов/типов несколько — показываем кратко
            return string.Join(", ", distinct.Take(3)) + (distinct.Count > 3 ? "…" : "");
        }

        static string NormalizeLine(string s) =>
            string.IsNullOrWhiteSpace(s) ? "(без линии)" : s.Trim();

        private static string SafeStatus(string raw)
        {
            if (string.IsNullOrWhiteSpace(raw)) return "Не определён";
            // Нормализуем под вашу легенду, если нужно
            return raw.Trim();
        }

        /// <summary>
        /// Пытается получить «Линия» из TechnicalInfo.
        /// Подставьте своё поле, если оно у вас другое.
        /// </summary>
        private static string ResolveLineName(TechnicalInfo info)
        {
            if (info == null) return string.Empty;

            // Популярные варианты имён свойства
            var type = typeof(TechnicalInfo);
            var pi = type.GetProperty("Line") ?? type.GetProperty("LineName") ?? type.GetProperty("Tag");
            if (pi != null)
            {
                var val = pi.GetValue(info);
                return val?.ToString() ?? string.Empty;
            }
            return string.Empty;
        }


        private DelegateCommand createReportPassports;
        public ICommand CreateReportPassports
        {
            get
            {
                if (createReportPassports == null)
                {
                    createReportPassports = new DelegateCommand(async () => await PerformCreateReportPassportsAsync());
                }

                return createReportPassports;
            }
        }

        private async Task PerformCreateReportPassportsAsync()
        {
            // Аккумуляторы итогов (для PDF и контроля)
            double totalTonnageT = 0.0;   // т
            double totalLengthM = 0.0;   // м
            double totalWeldInch = 0.0;   // inch

            // 1) Проверка данных
            if (TestData == null || TestData.Items == null || TestData.Items.Count == 0)
            {
                MessageBox.Show("Нет данных для отчёта по Паспортам.", "Пусто",
                    MessageBoxButton.OK, MessageBoxImage.Information);
                return;
            }

            // 2) Диалог сохранения
            var sfd = new SaveFileDialog
            {
                Title = "Сохранить отчёт 'Паспорта'",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx|PDF Document (*.pdf)|*.pdf",
                FileName = $"Паспорта_{DateTime.Now:yyyyMMdd_HHmm}"
            };
            if (sfd.ShowDialog() != true)
                return;

            // 3) Окно прогресса
            var dlg = new ProgressDialog { Owner = Application.Current.MainWindow };
            dlg.SetIndeterminate("Подготовка данных...");
            dlg.Show();

            IProgress<ReportProgress> progress =
                new System.Progress<ReportProgress>(p => dlg.SetProgress(p.Phase, p.Current, p.Total));

            try
            {
                await Task.Run(() =>
                {
                    var token = dlg.Token;
                    token.ThrowIfCancellationRequested();

                    // ---------- ФИЛЬТР ПО НОВОЙ ЛЕГЕНДЕ (иерархия) ----------
                    // Разрешённые GUID только из ВИДИМЫХ секций (детей).
                    // Включаем фильтр ТОЛЬКО если выбрано "не всё": т.е. есть видимые, но их меньше, чем всех.
                    HashSet<Guid> allowed = null;
                    if (PP_LegendTree != null && PP_LegendTree.Count > 0)
                    {
                        var visibleIds = PP_LegendTree
                            .SelectMany(p => p.Children)
                            .Where(c => c.IsVisible)
                            .SelectMany(c => c.ElementIds ?? Array.Empty<IModelElementId>())
                            .OfType<ForColorModelElementId>()
                            .Select(x => x.ElementId)
                            .ToList();

                        var allIds = PP_LegendTree
                            .SelectMany(p => p.Children)
                            .SelectMany(c => c.ElementIds ?? Array.Empty<IModelElementId>())
                            .OfType<ForColorModelElementId>()
                            .Select(x => x.ElementId)
                            .ToList();

                        if (visibleIds.Count > 0 && visibleIds.Count < allIds.Count)
                            allowed = new HashSet<Guid>(visibleIds);
                    }

                    // ---------- ПОДГОТОВКА ДАННЫХ БЕЗ ДУБЛЕЙ ----------
                    // Берём ровно по одному разу каждый уникальный элемент (Guid),
                    // при этом уважаем фильтр "allowed" (если он активен).
                    var idsUniverse = TestData.Items.Keys; // уникальные GUID
                    var idsToUse = (allowed == null) ? idsUniverse : idsUniverse.Where(g => allowed.Contains(g));

                    // Преобразуем в «строчки» с уже вычисленными ключами
                    // Паспорт/Секция нормализуем так же, как в легенде: "ПАСПОРТ | СЕКЦИЯ"
                    Func<string, string> Norm = s => string.IsNullOrWhiteSpace(s) ? "НЕ ОПРЕДЕЛЁН" : s.Trim().ToUpperInvariant();
                    Func<string, string> Pretty = s => s == "НЕ ОПРЕДЕЛЁН" ? "Не определён" : s;
                    const string SectionSeparator = " | ";

                    var rowsPerElement = new List<dynamic>(capacity: TestData.Items.Count);

                    foreach (var id in idsToUse)
                    {
                        token.ThrowIfCancellationRequested();

                        TechnicalInfo info;
                        if (!TestData.Items.TryGetValue(id, out info) || info == null)
                            continue;

                        string pass = Pretty(Norm(info.PassportId));
                        string sect = Pretty(Norm(info.Section == null ? null : info.Section.ToString()));
                        string passportSection = pass + SectionSeparator + sect;

                        string line = NormalizeLine(ResolveLineName(info));

                        double len = info.TotalTpLengthMeter ?? 0.0;
                        double ton_T = (info.TotalTpWeightKg ?? 0.0) / 1000.0; // кг -> т
                        double weld = info.TotalTpWdi ?? 0.0;

                        rowsPerElement.Add(new
                        {
                            PassportSection = passportSection,
                            Line = line,
                            LengthM = len,
                            TonnageT = ton_T,
                            WeldInch = weld
                        });

                        // Периодически обновляем прогресс
                        if ((rowsPerElement.Count & 255) == 0)
                            progress.Report(new ReportProgress { Phase = "Подготовка данных...", Current = rowsPerElement.Count, Total = TestData.Items.Count });
                    }

                    // ---------- ГРУППИРОВАНИЕ ДЛЯ ОТЧЁТА ----------
                    token.ThrowIfCancellationRequested();
                    progress.Report(new ReportProgress { Phase = "Агрегирование...", Current = 0, Total = 0 });

                    var grouped = rowsPerElement
                        .GroupBy(r => new { r.PassportSection, r.Line })
                        .OrderBy(g => g.Key.PassportSection)
                        .ThenBy(g => g.Key.Line)
                        .ToList();

                    // ---------- ФОРМИРОВАНИЕ XLSX / PDF ----------
                    progress.Report(new ReportProgress { Phase = "Формирование файла...", Current = 0, Total = grouped.Count });

                    var wb = new XLWorkbook();
                    var ws = wb.Worksheets.Add("Паспорта");

                    // Заголовки
                    ws.Cell(1, 1).Value = "Паспорт";           // «Паспорт | Секция»
                    ws.Cell(1, 2).Value = "Линия";
                    ws.Cell(1, 3).Value = "Тоннаж, т";
                    ws.Cell(1, 4).Value = "Протяженность, м";
                    ws.Cell(1, 5).Value = "Сварной шов, inch";

                    int row = 2;
                    int processed = 0;

                    foreach (var g in grouped)
                    {
                        token.ThrowIfCancellationRequested();

                        string passportSection = g.Key.PassportSection;
                        string line = g.Key.Line;

                        double tonnageT = g.Sum(x => (double)x.TonnageT);
                        double lengthM = g.Sum(x => (double)x.LengthM);
                        double weldInch = g.Sum(x => (double)x.WeldInch);

                        ws.Cell(row, 1).Value = passportSection;
                        ws.Cell(row, 2).Value = line;
                        ws.Cell(row, 3).Value = tonnageT;
                        ws.Cell(row, 4).Value = lengthM;
                        ws.Cell(row, 5).Value = weldInch;

                        totalTonnageT += tonnageT;
                        totalLengthM += lengthM;
                        totalWeldInch += weldInch;

                        row++;
                        processed++;
                        if ((processed & 63) == 0)
                            progress.Report(new ReportProgress { Phase = "Формирование файла...", Current = processed, Total = grouped.Count });
                    }

                    // Оформление таблицы
                    int lastDataRow = Math.Max(1, row - 1);
                    var rng = ws.Range(1, 1, lastDataRow, 5);
                    var table = rng.CreateTable();
                    table.Theme = XLTableTheme.TableStyleLight9;

                    ws.Column(3).Style.NumberFormat.Format = "#,##0.000";
                    ws.Column(4).Style.NumberFormat.Format = "#,##0.00";
                    ws.Column(5).Style.NumberFormat.Format = "#,##0.00";

                    table.ShowTotalsRow = true;
                    table.Field(0).TotalsRowLabel = "ИТОГО:";
                    table.Field(2).TotalsRowLabel = totalTonnageT.ToString("N2");
                    table.Field(3).TotalsRowLabel = totalLengthM.ToString("N2");
                    table.Field(4).TotalsRowLabel = totalWeldInch.ToString("N2");

                    ws.SheetView.FreezeRows(1);
                    ws.Columns(1, 5).AdjustToContents();

                    // Сохранение
                    progress.Report(new ReportProgress { Phase = "Сохранение файла...", Current = 0, Total = 0 });

                    string ext = System.IO.Path.GetExtension(sfd.FileName).ToLowerInvariant();

                    if (ext == ".xlsx")
                    {
                        wb.SaveAs(sfd.FileName);
                    }
                    else if (ext == ".pdf")
                    {
                        var headers = new[]
                        {
                    "Паспорт",           // «Паспорт | Секция»
                    "Линия",
                    "Тоннаж, т",
                    "Протяженность, м",
                    "Сварной шов, inch"
                };

                        var rows = new List<string[]>(capacity: grouped.Count + 2);
                        foreach (var g in grouped)
                        {
                            string passportSection = g.Key.PassportSection;
                            string line = g.Key.Line;

                            double tonnageT = g.Sum(x => (double)x.TonnageT);
                            double lengthM = g.Sum(x => (double)x.LengthM);
                            double weldInch = g.Sum(x => (double)x.WeldInch);

                            rows.Add(new[]
                            {
                        passportSection,
                        line,
                        tonnageT.ToString("#,##0.000"),
                        lengthM.ToString("#,##0.00"),
                        weldInch.ToString("#,##0.00")
                    });
                        }

                        rows.Add(new[]
                        {
                    "ИТОГО:", "",
                    totalTonnageT.ToString("#,##0.000"),
                    totalLengthM.ToString("#,##0.00"),
                    totalWeldInch.ToString("#,##0.00")
                });

                        GeneratePdfReport(
                            filePath: sfd.FileName,
                            title: "Отчёт по паспортам трубопроводов" + " " + ProjectName,
                            createdAt: DateTime.Now,
                            headers: headers,
                            rows: rows
                        );
                    }
                }, dlg.Token);

                dlg.Close();
                MessageBox.Show("Отчёт успешно сформирован.", "Готово", MessageBoxButton.OK, MessageBoxImage.Information);

                if (System.IO.Path.GetExtension(sfd.FileName).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = sfd.FileName,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Не удалось открыть PDF автоматически:\n{ex.Message}", "Предупреждение",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                dlg.Close();
                MessageBox.Show("Формирование отчёта отменено.", "Отмена", MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            catch (Exception ex)
            {
                dlg.Close();
                MessageBox.Show($"Ошибка при формировании отчёта:\n{ex.Message}", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        void GeneratePdfReport(string filePath, string title, DateTime createdAt, string[] headers, List<string[]> rows)
        {
            var builder = new PdfReportBuilder(title, createdAt, headers, rows);
            builder.GeneratePdf(filePath);
        }


        #endregion

        #region Паспортизация трубопроводов

        private double PP_totalTonnage;
        public double PP_TotalTonnage
        {
            get => PP_totalTonnage;
            set
            {
                if (Math.Abs(PP_totalTonnage - value) < 0.01) return;
                PP_totalTonnage = value;
                OnPropertyChanged();
            }
        }

        private double PP_totalLength;
        public double PP_TotalLength
        {
            get => PP_totalLength;
            set
            {
                if (Math.Abs(PP_totalLength - value) < 0.01) return;
                PP_totalLength = value;
                OnPropertyChanged();
            }
        }

        private double PP_totalWeld;
        public double PP_TotalWeld
        {
            get => PP_totalWeld;
            set
            {
                if (Math.Abs(PP_totalWeld - value) < 0.01) return;
                PP_totalWeld = value;
                OnPropertyChanged();
            }
        }

        private static string Norm(string s)
        {
            return string.IsNullOrWhiteSpace(s) ? "НЕ ОПРЕДЕЛЁН" : s.Trim().ToUpperInvariant();
        }

        private static string Pretty(string s)
        {
            return s == "НЕ ОПРЕДЕЛЁН" ? "Не определён" : s;
        }

        // Числовые секции сортируем первыми: "0|0000000123", затем строковые: "1|ABC"
        private static string SectionSortKey(string sectionKey)
        {
            if (string.IsNullOrWhiteSpace(sectionKey)) return "0|0000000000";
            sectionKey = sectionKey.Trim();

            int n;
            if (int.TryParse(sectionKey, NumberStyles.Integer, CultureInfo.InvariantCulture, out n))
                return "0|" + n.ToString("0000000000", CultureInfo.InvariantCulture);

            return "1|" + sectionKey.ToUpperInvariant();
        }

        // ====== ПОСТРОЕНИЕ ДЕРЕВА ЛЕГЕНДЫ ======
        private void PP_MakeLegend()
        {
            OnClearAll();

            // Если где-то ещё используется плоская легенда — очистим её тоже
            PP_Legend.Clear();
            PP_LegendTree.Clear();

            // Группировка: Паспорт -> Секция -> Элементы
            var byPassport = TestData.Items.Values
                .GroupBy(info => Norm(info.PassportId))
                .OrderBy(g => g.Key, StringComparer.Ordinal);

            foreach (var passportGroup in byPassport)
            {
                var passportKey = passportGroup.Key;
                // цвет родителя (только для кружка в легенде)
                var parentColor = GetNextUniqueColor();
                var parentWpfColor = System.Windows.Media.Color.FromArgb(255, parentColor.R, parentColor.G, parentColor.B);
                var parentBrush = new System.Windows.Media.SolidColorBrush(parentWpfColor);

                var parent = new PP_PassportLegendNode(Pretty(passportKey), parentBrush);


                var bySection = passportGroup
                    .GroupBy(info => Norm(info.Section == null ? null : info.Section.ToString()))
                    .OrderBy(sg => SectionSortKey(sg.Key));

                foreach (var sectionGroup in bySection)
                {
                    var sectionKey = sectionGroup.Key;

                    var elementIds = sectionGroup
                        .SelectMany(info => info.ModelElementIds.Cast<IModelElementId>())
                        .GroupBy(e => e.ElementId)
                        .Select(gx => gx.First())
                        .ToArray();

                    var color = GetNextUniqueColor();
                    var wpfColor = System.Windows.Media.Color.FromArgb(255, color.R, color.G, color.B);
                    var brush = new System.Windows.Media.SolidColorBrush(wpfColor);

                    // Подпись: «Паспорт | Секция»
                    var childLabel = Pretty(passportKey) + SectionSeparator + Pretty(sectionKey);

                    var vm = new PP_LegendItem(
                        childLabel,
                        brush,
                        elementIds,
                        this,
                        _modelViewer,
                        false);

                    vm.IsVisible = false; // по умолчанию не включаем подсветку
                    parent.Children.Add(vm);
                }

                parent.HookChildren();
                PP_LegendTree.Add(parent);
            }

            PP_RecalculatePhysicalSumms();
        }

        // ====== ПЕРЕСЧЁТ СУММ ПО ВИДИМЫМ СЕКЦИЯМ (без дублей) ======
        public void PP_RecalculatePhysicalSumms()
        {
            double totalLength = 0.0;
            double totalTonnage = 0.0;
            double totalWELD = 0.0;

            // 1) Собираем все уникальные GUID элементов из ВИДИМЫХ секций
            var visibleGuids = new HashSet<Guid>();

            foreach (var passport in PP_LegendTree)
            {
                foreach (var item in passport.Children)
                {
                    if (!item.IsVisible) continue;

                    // ElementIds уже могли быть "очищены" от дублей, но HashSet гарантирует уникальность на всё дерево
                    foreach (var modelId in item.ElementIds)
                    {
                        var fcId = modelId as ForColorModelElementId;
                        if (fcId == null) continue;

                        visibleGuids.Add(fcId.ElementId);
                    }
                }
            }

            // 2) ОДИН раз проходим по уникальным GUID и суммируем
            foreach (var id in visibleGuids)
            {
                TechnicalInfo info;
                if (!TestData.Items.TryGetValue(id, out info)) continue;

                totalLength += info.TotalTpLengthMeter ?? 0.0;
                totalTonnage += (info.TotalTpWeightKg ?? 0.0) / 1000.0; // кг → т
                totalWELD += info.TotalTpWdi ?? 0.0;
            }

            PP_TotalLength = totalLength;
            PP_TotalTonnage = totalTonnage;
            PP_TotalWeld = totalWELD;
        }



        public static Dictionary<string, List<IModelElementId>> BuildPassportsForLegend(TPData data)
        {
            return data.Items
                .Values
                .GroupBy(info => MakeKey(info.PassportId, info.Section?.ToString()))
                .ToDictionary(
                    grp => grp.Key,
                    grp => grp.SelectMany(info => info.ModelElementIds.Select(id => (IModelElementId)id)).ToList()
                );
        }

        #region Кнопки для быстрого выделения
        // ====== КНОПКИ МАССОВЫХ ДЕЙСТВИЙ (работают по дочерним узлам) ======
        private DelegateCommand selectAll_Passports;
        public ICommand SelectAll_Passports
        {
            get
            {
                if (selectAll_Passports == null)
                    selectAll_Passports = new DelegateCommand(PerformSelectAll_Passports);
                return selectAll_Passports;
            }
        }

        private void PerformSelectAll_Passports()
        {
            foreach (var p in PP_LegendTree)
                foreach (var c in p.Children)
                    c.IsVisible = true;
        }

        private DelegateCommand clearAll_Passports;
        public ICommand ClearAll_Passports
        {
            get
            {
                if (clearAll_Passports == null)
                    clearAll_Passports = new DelegateCommand(PerformClearAll_Passports);
                return clearAll_Passports;
            }
        }

        public void PerformClearAll_Passports()
        {
            foreach (var p in PP_LegendTree)
                p.IsChecked = false;
        }

        private DelegateCommand hideUnhighlighted_Passports;
        public ICommand HideUnhighlighted_Passports
        {
            get
            {
                if (hideUnhighlighted_Passports == null)
                    hideUnhighlighted_Passports = new DelegateCommand(PerformHideUnhighlighted_Passports);
                return hideUnhighlighted_Passports;
            }
        }

        private void PerformHideUnhighlighted_Passports()
        {
            foreach (var p in PP_LegendTree)
            {
                foreach (var c in p.Children)
                {
                    c.HideU(true);
                    if (!c.IsVisible)
                        _modelViewer.Hide(c.ElementIds.GroupBy(e => e.ElementId).Select(g => g.First()).ToArray());
                }
            }
        }

        private DelegateCommand restoreVisibility_Passsports;
        public ICommand RestoreVisibility_Passsports
        {
            get
            {
                if (restoreVisibility_Passsports == null)
                    restoreVisibility_Passsports = new DelegateCommand(PerformRestoreVisibility_Passsports);
                return restoreVisibility_Passsports;
            }
        }

        public void PerformRestoreVisibility_Passsports()
        {
            foreach (var p in PP_LegendTree)
            {
                foreach (var c in p.Children)
                {
                    c.HideU(false);
                    if (!c.IsVisible)
                        _modelViewer.Show(c.ElementIds.GroupBy(e => e.ElementId).Select(g => g.First()).ToArray());
                }
            }
        }
        #endregion

        #region Вспомогательные штуки
        private const string SectionSeparator = " | ";

        private static string NormalizeToken(string s)
        {
            if (string.IsNullOrWhiteSpace(s)) return "НЕ ОПРЕДЕЛЁН";
            return s.Trim().ToUpperInvariant();
        }

        private static string NormalizePassportKey(string passportId) => NormalizeToken(passportId);
        private static string NormalizeSectionKey(string section) => NormalizeToken(section);

        // Собираем нормализованный ключ «ПАСПОРТ | СЕКЦИЯ»
        private static string MakeKey(string passportId, string section)
            => $"{NormalizePassportKey(passportId)}{SectionSeparator}{NormalizeSectionKey(section)}";

        // Парсим ключ на части + пытаемся получить числовой номер секции
        private static (string passportKey, string sectionKey, int? sectionNum) SplitKey(string key)
        {
            var parts = key.Split(new[] { SectionSeparator }, StringSplitOptions.None);
            var p = parts.Length > 0 ? parts[0] : "НЕ ОПРЕДЕЛЁН";
            var s = parts.Length > 1 ? parts[1] : "НЕ ОПРЕДЕЛЁН";

            if (int.TryParse(s, NumberStyles.Integer, CultureInfo.InvariantCulture, out var num))
                return (p, s, num);

            return (p, s, null);
        }

        // Человекочитаемая подпись для легенды
        private static string PrettyFromKey(string key)
        {
            var (p, s, _) = SplitKey(key);
            string PrettyToken(string t) => t == "НЕ ОПРЕДЕЛЁН" ? "Не определён" : t;
            return $"{PrettyToken(p)}{SectionSeparator}{PrettyToken(s)}";
        }

        private List<Ascon.Pilot.SDK.IDataObject> dataObjects;

        public void get_dataobjects(List<Ascon.Pilot.SDK.IDataObject> objects)
        {
            dataObjects = objects;

            if (dataObjects != null)
            {
                string message = "";

                foreach (var element in dataObjects)
                {
                    var section = element.DisplayName.Split('_')[0].Split('-')[0];
                    message += section + ": " + element.Id.ToString() + "\n";

                    partMap_guid_string[element.Id] = section;
                    partMap_string_guid[section] = element.Id;
                }

                //MessageBox.Show(message);
            }
            else
            {
                MessageBox.Show($"dataObjects is null");
            }
        }





        #endregion

        #endregion


        #region Статусы подсистем

        private DelegateCommand createReportSubsystemsStatus;
        public ICommand CreateReportSubsystemsStatus
        {
            get
            {
                if (createReportSubsystemsStatus == null)
                {
                    createReportSubsystemsStatus = new DelegateCommand(PerformCreateReportSubsystemsStatus);
                }

                return createReportSubsystemsStatus;
            }
        }

        /// <summary>
        /// Команда на создание отчёта по подсистемам (Excel / PDF).
        /// </summary>
        private async void PerformCreateReportSubsystemsStatus()
        {
            if (Systems == null)
            {
                MessageBox.Show("Сначала загрузите данные в первой вкладке!");
                return;
            }

            //var statusReader = new SubsystemsStatusReader(_filePath + "Highlighter\\");
            //var lineStatuses = statusReader.ReadLineStatuses();

            //if (lineStatuses == null || lineStatuses.Count == 0)
            //{
            //    MessageBox.Show("Нет данных по линиям для формирования отчёта по подсистемам.", "Ошибка",
            //        MessageBoxButton.OK, MessageBoxImage.Error);
            //    return;
            //}

            //var mapper = new SubsystemsStatusMapper(Systems);

            var reader = new SubsystemsStatusReader(_filePath + "Highlighter\\");

            // Статусы линий: приоритеты подмешиваются из "Приоритеты (по линиям).xlsx"
            var lineStatuses = reader.ReadLineStatusesWithExternalPriorities();

            // Новая структура подсистем: "Разбиение на подсистемы (по линиям).xlsx"
            var subsystemToLines = reader.ReadSubsystemsStructureFromLines();

            var mapper = new SubsystemsStatusMapper(Systems);

            var newMapper = mapper.BuildRemappedMapperFromLineMapping(subsystemToLines);

            await PerformCreateReportSubsystemsStatusAsync(newMapper, lineStatuses);
        }

        /// <summary>
        /// Отладочный метод: случайные N линий из исходных данных.
        /// </summary>
        public static string GetRandomLinesSummary(
            Dictionary<string, SubsystemsStatusReader.LineStatus> data,
            int count = 10)
        {
            if (data == null || data.Count == 0)
                return "Нет данных по линиям.";

            var rnd = new Random();
            var lines = data.Values.ToList();

            int take = Math.Min(count, lines.Count);

            var selected = lines.OrderBy(x => rnd.Next()).Take(take).ToList();

            var sb = new System.Text.StringBuilder();
            sb.AppendLine($"Случайные {take} линий:");

            foreach (var line in selected)
            {
                sb.AppendLine(
                    $"{line.LineName} | TP: {line.TestPackageName} | " +
                    $"Приоритет: {line.Priority} | " +
                    $"Сварка: {line.WeldingPercent:0.#}% | " +
                    $"НК: {line.NdtPercent:0.#}% | "
                );
            }

            return sb.ToString();
        }

        /// <summary>
        /// Собственно формирование Excel / PDF отчёта по подсистемам на основе SubsystemsStatusMapper.
        /// </summary>
        private async Task PerformCreateReportSubsystemsStatusAsync(
            SubsystemsStatusMapper mapper,
            Dictionary<string, SubsystemsStatusReader.LineStatus> lineStatuses)
        {
            if (mapper == null)
            {
                MessageBox.Show("Mapper для подсистем не передан.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            if (lineStatuses == null || lineStatuses.Count == 0)
            {
                MessageBox.Show("Нет данных по линиям для формирования отчёта по подсистемам.", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            // 1) Диалог выбора файла — до запуска фоновой задачи
            var sfd = new SaveFileDialog
            {
                Title = "Сохранить отчёт по подсистемам",
                Filter = "Excel Workbook (*.xlsx)|*.xlsx|PDF Document (*.pdf)|*.pdf",
                FileName = $"Статус_подсистем_{DateTime.Now:yyyyMMdd_HHmm}"
            };
            if (sfd.ShowDialog() != true)
                return;

            // 2) Окно прогресса
            var dlg = new ProgressDialog { Owner = Application.Current.MainWindow };
            dlg.SetIndeterminate("Подготовка данных...");
            dlg.Show();

            IProgress<ReportProgress> progress =
                new System.Progress<ReportProgress>(p =>
                {
                    dlg.SetProgress(p.Phase, p.Current, p.Total);
                });

            try
            {
                await Task.Run(() =>
                {
                    var token = dlg.Token;
                    token.ThrowIfCancellationRequested();

                    progress.Report(new ReportProgress
                    {
                        Phase = "Агрегация данных по подсистемам...",
                        Current = 0,
                        Total = 0
                    });

                    // --- Стадия 1: агрегируем по подсистемам через SubsystemsStatusMapper ---
                    var subsystemsData = mapper.BuildSubsystemStatuses(lineStatuses);

                    // Плоский список сводок, чтобы удобнее и для Excel, и для PDF
                    var flat = new List<SubsystemsStatusMapper.SubsystemStatusSummary>();
                    foreach (var sysKv in subsystemsData)
                    {
                        foreach (var subsKv in sysKv.Value)
                            flat.Add(subsKv.Value);
                    }

                    if (flat.Count == 0)
                    {
                        throw new InvalidOperationException("Не удалось собрать данные по подсистемам (пустой результат).");
                    }

                    // Локальный хелпер: посчитать по подсистеме:
                    // - всего линий
                    // - принято линий
                    // - всего тест-пакетов
                    // - принято тест-пакетов
                    (int totalLines, int acceptedLines, int totalTp, int acceptedTp) ComputeSubsystemStats(
                        string systemName,
                        string subsystemName)
                    {
                        int totalLines = 0;
                        int acceptedLines = 0;
                        int totalTp = 0;
                        int acceptedTp = 0;

                        if (string.IsNullOrWhiteSpace(systemName) || string.IsNullOrWhiteSpace(subsystemName))
                            return (0, 0, 0, 0);

                        if (!mapper.Systems.TryGetValue(systemName, out var system) ||
                            system.Subsystems == null || system.Subsystems.Count == 0)
                            return (0, 0, 0, 0);

                        var subsystem = system.Subsystems
                            .FirstOrDefault(s => string.Equals(s.Name, subsystemName, StringComparison.OrdinalIgnoreCase));

                        if (subsystem == null || subsystem.PipeLines == null || subsystem.PipeLines.Count == 0)
                            return (0, 0, 0, 0);

                        var tpTotalSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                        var tpAcceptedSet = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

                        foreach (var pipeLine in subsystem.PipeLines)
                        {
                            if (pipeLine == null || string.IsNullOrWhiteSpace(pipeLine.Name))
                                continue;

                            var normName = pipeLine.Name.Trim();
                            if (string.IsNullOrEmpty(normName))
                                continue;

                            totalLines++;

                            if (!lineStatuses.TryGetValue(normName, out var ls) || ls == null)
                                continue;

                            if (!string.IsNullOrWhiteSpace(ls.TestPackageName))
                            {
                                tpTotalSet.Add(ls.TestPackageName);

                                if (ls.TestsCompleted)
                                    tpAcceptedSet.Add(ls.TestPackageName);
                            }

                            if (ls.TestsCompleted)
                                acceptedLines++;
                        }

                        totalTp = tpTotalSet.Count;
                        acceptedTp = tpAcceptedSet.Count;

                        return (totalLines, acceptedLines, totalTp, acceptedTp);
                    }

                    // --- Стадия 2: формирование Excel ---
                    progress.Report(new ReportProgress
                    {
                        Phase = "Формирование Excel...",
                        Current = 0,
                        Total = flat.Count
                    });

                    var wb = new XLWorkbook();
                    var ws = wb.Worksheets.Add("Подсистемы");

                    // Заголовки
                    ws.Cell(1, 1).Value = "Система";
                    ws.Cell(1, 2).Value = "Подсистема";
                    ws.Cell(1, 3).Value = "Приоритеты";
                    ws.Cell(1, 4).Value = "Тест-пакеты (всего/принято)";
                    ws.Cell(1, 5).Value = "Средняя готовность по сварке, %";
                    ws.Cell(1, 6).Value = "Средняя готовность по НК, %";
                    ws.Cell(1, 7).Value = "Кол-во линий в подсистеме (всего/принято)";
                    ws.Cell(1, 8).Value = "% готовности подсистемы";

                    int rowIdx = 2;
                    int current = 0;
                    var inv = CultureInfo.InvariantCulture;

                    const string NoLinesText = "Линий из этой подсистемы\nнет в шаблоне";

                    foreach (var s in flat.OrderBy(x => x.SystemName).ThenBy(x => x.SubsystemName))
                    {
                        token.ThrowIfCancellationRequested();

                        // Базовые поля

                        // ФОКУС
                        string sysname = "";
                        if (s.SystemName.Contains("01-"))
                        {
                            var parts = s.SubsystemName.Split('-');
                            sysname = parts.Count() > 1 ? "01-" + parts[1] : s.SystemName;

                            //MessageBox.Show($"Поменял systemName ({s.SystemName}) для подсистемы {s.SubsystemName}");
                        }

                        ws.Cell(rowIdx, 1).Value = sysname != "" ? sysname : s.SystemName;
                        ws.Cell(rowIdx, 2).Value = s.SubsystemName ?? string.Empty;

                        ws.Cell(rowIdx, 3).Value = s.PrioritiesText ?? string.Empty;

                        // Считаем статистику по линиям и тест-пакетам
                        var (totalLines, acceptedLines, totalTp, acceptedTp) =
                            ComputeSubsystemStats(s.SystemName, s.SubsystemName);

                        // Тест-пакеты: всего/принято
                        ws.Cell(rowIdx, 4).Value = $"{totalTp}/{acceptedTp}";

                        // Средние проценты (как и раньше)
                        ws.Cell(rowIdx, 5).Value = s.AvgWeldingPercent;
                        ws.Cell(rowIdx, 6).Value = s.AvgNdtPercent;

                        // Кол-во линий: всего/принято
                        var cell7 = ws.Cell(rowIdx, 7);
                        if (totalLines == 0)
                        {
                            cell7.Value = NoLinesText;
                            cell7.Style.Alignment.WrapText = true;
                        }
                        else
                        {
                            cell7.Value = $"{totalLines}/{acceptedLines}";
                        }

                        // % готовности подсистемы = принято_линий / всего_линий * 100
                        double readinessPercent = 0.0;
                        if (totalLines > 0)
                            readinessPercent = 100.0 * acceptedLines / totalLines;

                        ws.Cell(rowIdx, 8).Value = readinessPercent;

                        rowIdx++;
                        current++;
                        if ((current & 31) == 0 || current == flat.Count)
                        {
                            progress.Report(new ReportProgress
                            {
                                Phase = "Формирование Excel...",
                                Current = current,
                                Total = flat.Count
                            });
                        }
                    }

                    // --- Стадия 3: оформление Excel ---
                    progress.Report(new ReportProgress
                    {
                        Phase = "Оформление Excel...",
                        Current = 0,
                        Total = 0
                    });

                    int lastDataRow = rowIdx - 1;
                    var rng = ws.Range(1, 1, lastDataRow, 8);
                    var table = rng.CreateTable();
                    table.Theme = XLTableTheme.TableStyleLight9;

                    // Форматы числовых колонок
                    ws.Column(5).Style.NumberFormat.Format = "#,##0.0"; // сварка
                    ws.Column(6).Style.NumberFormat.Format = "#,##0.0"; // НК
                    ws.Column(8).Style.NumberFormat.Format = "#,##0.0"; // % готовности подсистемы

                    ws.SheetView.FreezeRows(1);
                    ws.Columns(1, 8).AdjustToContents();

                    // --- Стадия 4: сохранение ---
                    progress.Report(new ReportProgress
                    {
                        Phase = "Сохранение файла...",
                        Current = 0,
                        Total = 0
                    });

                    string ext = System.IO.Path.GetExtension(sfd.FileName).ToLowerInvariant();
                    if (ext == ".xlsx")
                    {
                        wb.SaveAs(sfd.FileName);
                    }
                    else if (ext == ".pdf")
                    {
                        // Готовим данные для PdfReportBuilder
                        var headers = new[]
                        {
                    "Система",
                    "Подсистема",
                    "Приоритеты",
                    "Тест-пакеты (всего/принято)",
                    "Средняя готовность по сварке, %",
                    "Средняя готовность по НК, %",
                    "Кол-во линий в подсистеме (всего/принято)",
                    "% готовности подсистемы"
                };

                        var rows = new List<string[]>();

                        foreach (var s in flat.OrderBy(x => x.SystemName).ThenBy(x => x.SubsystemName))
                        {
                            var (totalLines, acceptedLines, totalTp, acceptedTp) =
                                ComputeSubsystemStats(s.SystemName, s.SubsystemName);

                            string avgWelding = s.AvgWeldingPercent.ToString("0.0", inv);
                            string avgNdt = s.AvgNdtPercent.ToString("0.0", inv);

                            string linesText = totalLines == 0
                                ? NoLinesText
                                : $"{totalLines}/{acceptedLines}";

                            double readinessPercent = 0.0;
                            if (totalLines > 0)
                                readinessPercent = 100.0 * acceptedLines / totalLines;

                            string readinessText = readinessPercent.ToString("0.0", inv);

                            // ФОКУС
                            string sysname = "";
                            if (s.SystemName.Contains("01-"))
                            {
                                var parts = s.SubsystemName.Split('-');
                                sysname = parts.Count() > 1 ? "01-" + parts[1] : s.SystemName;

                                //MessageBox.Show($"Поменял systemName ({s.SystemName}) для подсистемы {s.SubsystemName}");
                            }

                            rows.Add(new[]
                            {
                        sysname != "" ? sysname : s.SystemName,
                        s.SubsystemName ?? string.Empty,
                        s.PrioritiesText ?? string.Empty,
                        $"{totalTp}/{acceptedTp}",
                        avgWelding,
                        avgNdt,
                        linesText,
                        readinessText
                    });
                        }

                        var pdf = new PdfReportBuilder(
                            "Отчёт по статусу подсистем" + " " + ProjectName,
                            DateTime.Now,
                            headers,
                            rows);

                        pdf.GeneratePdf(sfd.FileName);
                    }

                }, dlg.Token);

                dlg.Close();
                MessageBox.Show("Отчёт по подсистемам успешно сформирован.", "Готово",
                    MessageBoxButton.OK, MessageBoxImage.Information);

                // Если PDF — сразу пробуем открыть
                if (System.IO.Path.GetExtension(sfd.FileName).Equals(".pdf", StringComparison.OrdinalIgnoreCase))
                {
                    try
                    {
                        System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                        {
                            FileName = sfd.FileName,
                            UseShellExecute = true
                        });
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show($"Не удалось открыть файл автоматически:\n{ex.Message}", "Ошибка",
                            MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
            catch (OperationCanceledException)
            {
                dlg.Close();
                MessageBox.Show("Формирование отчёта отменено.", "Отмена",
                    MessageBoxButton.OK, MessageBoxImage.Exclamation);
            }
            catch (Exception ex)
            {
                dlg.Close();
                MessageBox.Show($"Ошибка при формировании отчёта по подсистемам:\n{ex.Message}", "Ошибка",
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }


        // --- Легенда «Статус подсистем»: Подсистема → Линии ---

        /// <summary>
        /// Дерево для легенды: корни — подсистемы, дети — линии.
        /// </summary>
        public ObservableCollection<SubsystemLegendNode> SubsystemsLegendTree { get; }
            = new ObservableCollection<SubsystemLegendNode>();

        /// <summary>
        /// Режим просмотра для кнопок:
        /// "Готовность по сварке", "Средняя готовность НК", "Испытания завершены".
        /// </summary>
        public enum SubsystemStatusViewMode
        {
            None,
            Welding,
            Ndt,
            Tests
        }

        private SubsystemStatusViewMode _currentSubsystemViewMode = SubsystemStatusViewMode.None;

        // нужно, чтобы легенда/элементы знали текущий режим при клике по чекбоксам
        public SubsystemStatusViewMode CurrentSubsystemViewMode => _currentSubsystemViewMode;
        public bool HideUnhighlightedMode => HideUnhighlighted_mode;

        // Поисковая строка именно для вкладки "Статус подсистем"
        private string _subsystemsSearchText;
        public string SubsystemsSearchText
        {
            get => _subsystemsSearchText;
            set
            {
                if (_subsystemsSearchText == value) return;
                _subsystemsSearchText = value;
                OnPropertyChanged();
            }
        }

        // Команды для поиска
        public DelegateCommand SubsystemsSearchCommand { get; }
        public DelegateCommand ClearSubsystemsSearchCommand { get; }

        // Команды переключения режимов покраски
        public DelegateCommand ShowSubsystemsWeldingCommand { get; }
        public DelegateCommand ShowSubsystemsNdtCommand { get; }
        public DelegateCommand ShowSubsystemsTestsCommand { get; }

        // Кнопки видимости / чекбоксов для этой вкладки
        public DelegateCommand HideUnhighlightedSubsystemsCommand { get; }
        public DelegateCommand RestoreVisibilitySubsystemsCommand { get; }

        public DelegateCommand ActivateAllSubsystemsCommand { get; }
        public DelegateCommand DeactivateAllSubsystemsCommand { get; }

        /// <summary>
        /// Инициализация кэша по подсистемам и легенды.
        /// Вызывается после загрузки TP-данных.
        /// </summary>
        private async Task InitSubsystemStatusesAsync()
        {
            if (Systems == null)
                return;

            try
            {
                //var reader = new SubsystemsStatusReader(_filePath + "Highlighter\\");
                //var lineStatuses = reader.ReadLineStatuses();

                //if (lineStatuses == null || lineStatuses.Count == 0)
                //    return;

                //var mapper = new SubsystemsStatusMapper(Systems);
                //var subsystemsData = mapper.BuildSubsystemStatuses(lineStatuses);

                var reader = new SubsystemsStatusReader(_filePath + "Highlighter\\");

                // 1. Статусы линий
                var lineStatuses = reader.ReadLineStatusesWithExternalPriorities();

                // 2. Разбиение подсистем по линиям из Excel
                var subsystemToLines = reader.ReadSubsystemsStructureFromLines();

                // 3. Мапер по старой структуре (источник элементов)
                var originalMapper = new SubsystemsStatusMapper(Systems);

                // 4. Строим новый мапер по Excel-разбиению
                var remappedMapper = originalMapper.BuildRemappedMapperFromLineMapping(subsystemToLines);

                // 5. Статусы подсистем по новой структуре
                var subsystems = remappedMapper.BuildSubsystemStatusesFromLineMapping(lineStatuses, subsystemToLines);

                // 6. Легенда по НОВОЙ структуре
                BuildSubsystemsLegend(remappedMapper, subsystems);

            }
            catch (Exception ex)
            {
                MessageBox.Show(
                    $"Ошибка при инициализации статусов подсистем:\n{ex.Message}",
                    "Ошибка",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error);
            }
        }

        /// <summary>
        /// Строит дерево SubsystemsLegendTree: корни — подсистемы (из SubsystemsStatusMapper),
        /// дети — линии из systemsSource (обычно remappedMapper.Systems).
        /// Все кружки и чекбоксы изначально "выключены" (серые, без подсветки).
        /// </summary>
        private void BuildSubsystemsLegend(
            SubsystemsStatusMapper mapper,
            Dictionary<string, Dictionary<string, SubsystemsStatusMapper.SubsystemStatusSummary>> subsystemsData)
        {
            SubsystemsLegendTree.Clear();

            if (mapper == null)
                return;
            if (subsystemsData == null || subsystemsData.Count == 0)
                return;

            var systemsSource = mapper.Systems;
            if (systemsSource == null || systemsSource.Count == 0)
                return;

            foreach (var sysKv in subsystemsData.OrderBy(k => k.Key))
            {
                var systemName = sysKv.Key;

                if (!systemsSource.TryGetValue(systemName, out var system) ||
                    system.Subsystems == null || system.Subsystems.Count == 0)
                    continue;

                var subsDict = sysKv.Value;
                if (subsDict == null || subsDict.Count == 0)
                    continue;

                foreach (var subsKv in subsDict.OrderBy(k => k.Key))
                {
                    var subsName = subsKv.Key;
                    var summary = subsKv.Value;

                    // Берём подсистему именно из НОВОЙ структуры mapper.Systems,
                    // которая уже собрана по Excel-разбиению.
                    var subsystem = system.Subsystems
                        .FirstOrDefault(s => string.Equals(s.Name, subsName, StringComparison.OrdinalIgnoreCase));

                    if (subsystem == null || subsystem.PipeLines == null || subsystem.PipeLines.Count == 0)
                        continue;

                    var node = new SubsystemLegendNode(systemName, subsName, this, ModelViewer);

                    // Линии подсистемы: имена и ElementId'ы берём из mapper.Systems,
                    // проценты — из summary (одни и те же для всех линий подсистемы).
                    foreach (var pipe in subsystem.PipeLines.OrderBy(p => p.Name))
                    {
                        if (pipe.Elements == null || pipe.Elements.Count == 0)
                            continue;

                        var elementIds = pipe.Elements
                            .Select(e => (IModelElementId)e.ColorId)
                            .ToArray();

                        var item = new SubsystemLegendItem(
                            pipe.Name,
                            elementIds,
                            summary.AvgWeldingPercent,
                            summary.AvgNdtPercent,
                            summary.TestsCompletedPercent,
                            this,
                            ModelViewer);

                        node.Children.Add(item);
                    }

                    if (node.Children.Count > 0)
                    {
                        node.RecalculateFromChildren();
                        node.SetGray(); // стартовое состояние: серые кружки, ничего не подсвечено
                        SubsystemsLegendTree.Add(node);
                    }
                }
            }
        }


        /// <summary>
        /// Полный градиент от красного (0 %) через жёлтый (50 %) к зелёному (100 %).
        /// </summary>
        public System.Windows.Media.Color GetSubsystemColor(double p01)
        {
            if (p01 < 0.0) p01 = 0.0;
            if (p01 > 1.0) p01 = 1.0;

            byte r, g;

            if (p01 <= 0.5)
            {
                // 0..0.5: красный -> жёлтый
                // r = 255, g: 0 -> 255
                double t = p01 / 0.5;       // 0..1
                r = 255;
                g = (byte)(255 * t);
            }
            else
            {
                // 0.5..1: жёлтый -> зелёный
                // r: 255 -> 0, g = 255
                double t = (p01 - 0.5) / 0.5; // 0..1
                r = (byte)(255 * (1.0 - t));
                g = 255;
            }

            return System.Windows.Media.Color.FromRgb(r, g, 0);
        }


        /// <summary>
        /// Обработчики для кнопок "Готовность по сварке / Средняя готовность НК / Испытания завершены".
        /// Меняют режим и перекрашивают легенду + элементы.
        /// </summary>
        private void OnShowSubsystemsWelding()
        {
            ApplySubsystemsViewMode(SubsystemStatusViewMode.Welding);
        }

        private void OnShowSubsystemsNdt()
        {
            ApplySubsystemsViewMode(SubsystemStatusViewMode.Ndt);
        }

        private void OnShowSubsystemsTests()
        {
            ApplySubsystemsViewMode(SubsystemStatusViewMode.Tests);
        }

        /// <summary>
        /// Применить выбранный режим просмотра к легенде и 3D-модели.
        /// Реальное окрашивание элементов делается внутри SubsystemLegendNode/SubsystemLegendItem.
        /// </summary>
        private void ApplySubsystemsViewMode(SubsystemStatusViewMode mode)
        {
            _currentSubsystemViewMode = mode;

            if (SubsystemsLegendTree == null || SubsystemsLegendTree.Count == 0)
                return;

            // Сначала всё сбрасываем в серый и очищаем подсветку в 3D
            foreach (var node in SubsystemsLegendTree)
                node.SetGray();

            if (mode == SubsystemStatusViewMode.None)
                return;

            // Потом применяем выбранный режим (учитывая состояние чекбоксов)
            foreach (var node in SubsystemsLegendTree)
                node.ApplyMode(mode);
        }

        // Поиск/сброс поиска используют уже существующий ApplyFilter
        private void OnSubsystemsSearch()
        {
            ApplyFilter(SubsystemsSearchText);
        }

        private void OnClearSubsystemsSearch()
        {
            SubsystemsSearchText = string.Empty;
            ApplyFilter(null);
        }

        // --- Геометрическая видимость для вкладки "Статус подсистем" ---

        private void OnHideUnhighlightedSubsystems()
        {
            HideUnhighlightedSubsystems();
        }

        private void OnRestoreVisibilitySubsystems()
        {
            HideUnhighlighted_mode = false;

            if (SubsystemsLegendTree == null || SubsystemsLegendTree.Count == 0 || ModelViewer == null)
                return;

            // Показать все элементы, участвующие в легенде
            var allIds = SubsystemsLegendTree
                .SelectMany(n => n.Children)
                .SelectMany(c => c.ElementIds ?? Array.Empty<IModelElementId>())
                .OfType<ForColorModelElementId>()
                .GroupBy(e => e.ElementId)
                .Select(g => (IModelElementId)g.First())
                .ToArray();

            if (allIds.Length > 0)
                ModelViewer.Show(allIds);
        }

        /// <summary>
        /// Включает режим HideUnhighlighted_mode и приводит геометрию в соответствие с чекбоксами.
        /// Всё, что не помечено IsVisible, скрывается.
        /// </summary>
        public void HideUnhighlightedSubsystems()
        {
            HideUnhighlighted_mode = true;

            if (SubsystemsLegendTree == null || SubsystemsLegendTree.Count == 0 || ModelViewer == null)
                return;

            foreach (var node in SubsystemsLegendTree)
            {
                foreach (var child in node.Children)
                {
                    var elements = child.ElementIds;
                    if (elements == null || elements.Length == 0)
                        continue;

                    if (child.IsVisible)
                    {
                        ModelViewer.Show(elements);
                    }
                    else
                    {
                        ModelViewer.Hide(elements);
                    }
                }
            }
        }

        /// <summary>
        /// Включить все чекбоксы (подсистемы и линии).
        /// </summary>
        private void OnActivateAllSubsystems()
        {
            if (SubsystemsLegendTree == null || SubsystemsLegendTree.Count == 0)
                return;

            foreach (var node in SubsystemsLegendTree)
            {
                node.IsChecked = true;
                foreach (var child in node.Children)
                    child.IsVisible = true;
            }

            // В режиме HideUnhighlighted_mode сразу показать всё
            if (HideUnhighlighted_mode && ModelViewer != null)
            {
                var allIds = SubsystemsLegendTree
                    .SelectMany(n => n.Children)
                    .SelectMany(c => c.ElementIds ?? Array.Empty<IModelElementId>())
                    .OfType<ForColorModelElementId>()
                    .GroupBy(e => e.ElementId)
                    .Select(g => (IModelElementId)g.First())
                    .ToArray();

                if (allIds.Length > 0)
                    ModelViewer.Show(allIds);
            }

            // Перекрасить по текущему режиму, если он выбран
            if (_currentSubsystemViewMode != SubsystemStatusViewMode.None)
            {
                foreach (var node in SubsystemsLegendTree)
                    node.ApplyMode(_currentSubsystemViewMode);
            }
        }

        /// <summary>
        /// Выключить все чекбоксы (подсистемы и линии).
        /// </summary>
        private void OnDeactivateAllSubsystems()
        {
            if (SubsystemsLegendTree == null || SubsystemsLegendTree.Count == 0)
                return;

            foreach (var node in SubsystemsLegendTree)
            {
                node.IsChecked = false;
                foreach (var child in node.Children)
                    child.IsVisible = false;
            }

            // Всех красим в серый и сбрасываем подсветку
            foreach (var node in SubsystemsLegendTree)
                node.SetGray();

            // В режиме HideUnhighlighted_mode – сразу прячем всю геометрию из легенды
            if (HideUnhighlighted_mode && ModelViewer != null)
            {
                var allIds = SubsystemsLegendTree
                    .SelectMany(n => n.Children)
                    .SelectMany(c => c.ElementIds ?? Array.Empty<IModelElementId>())
                    .OfType<ForColorModelElementId>()
                    .GroupBy(e => e.ElementId)
                    .Select(g => (IModelElementId)g.First())
                    .ToArray();

                if (allIds.Length > 0)
                    ModelViewer.Hide(allIds);
            }
        }

        #endregion




    }

    /// <summary>
    /// Родительский узел легенды: Подсистема (внутри – линии).
    /// </summary>
    public class SubsystemLegendNode : INotifyPropertyChanged
    {
        private readonly SystemsHighlighterControlViewModel _parentVm;
        private readonly IModelViewer _viewer;

        private Brush _brush = Brushes.Gray;
        private bool _isChecked;   // чекбокс у подсистемы

        public string SystemName { get; }
        public string SubsystemName { get; }

        public string Label => $"{SubsystemName} ({SystemName})";

        public ObservableCollection<SubsystemLegendItem> Children { get; } =
            new ObservableCollection<SubsystemLegendItem>();

        public double WeldingPercent { get; private set; }
        public double NdtPercent { get; private set; }
        public double TestsPercent { get; private set; }

        public Brush Brush
        {
            get => _brush;
            private set
            {
                if (Equals(_brush, value)) return;
                _brush = value;
                OnPropertyChanged(nameof(Brush));
            }
        }

        public bool IsChecked
        {
            get => _isChecked;
            set
            {
                if (_isChecked == value) return;
                _isChecked = value;
                OnPropertyChanged(nameof(IsChecked));

                if (!_isChecked)
                {
                    // отключаем подсистему: серый кружок, все дети снимают видимость
                    Brush = Brushes.Gray;
                    foreach (var child in Children)
                        child.IsVisible = false;
                }
                else
                {
                    // включаем подсистему: все дети видимы, если режим выбран — сразу красим
                    foreach (var child in Children)
                        child.IsVisible = true;

                    var mode = _parentVm.CurrentSubsystemViewMode;
                    if (mode != SystemsHighlighterControlViewModel.SubsystemStatusViewMode.None)
                        ApplyMode(mode);
                }
            }
        }

        public SubsystemLegendNode(
            string systemName,
            string subsystemName,
            SystemsHighlighterControlViewModel parentVm,
            IModelViewer viewer)
        {
            SystemName = systemName;
            SubsystemName = subsystemName;
            _parentVm = parentVm;
            _viewer = viewer;
        }

        public void RecalculateFromChildren()
        {
            if (Children.Count == 0)
            {
                WeldingPercent = NdtPercent = TestsPercent = 0;
                return;
            }

            WeldingPercent = Children.Average(c => c.WeldingPercent);
            NdtPercent = Children.Average(c => c.NdtPercent);
            TestsPercent = Children.Average(c => c.TestsPercent);
        }

        /// <summary>
        /// Применить режим к узлу и всем детям. Учитывает IsChecked и IsVisible.
        /// </summary>
        public void ApplyMode(SystemsHighlighterControlViewModel.SubsystemStatusViewMode mode)
        {
            double p = 0.0;
            switch (mode)
            {
                case SystemsHighlighterControlViewModel.SubsystemStatusViewMode.Welding:
                    p = WeldingPercent;
                    break;
                case SystemsHighlighterControlViewModel.SubsystemStatusViewMode.Ndt:
                    p = NdtPercent;
                    break;
                case SystemsHighlighterControlViewModel.SubsystemStatusViewMode.Tests:
                    p = TestsPercent;
                    break;
            }

            var color = _parentVm.GetSubsystemColor(p / 100.0);
            var brush = new SolidColorBrush(color);
            brush.Freeze();

            // если подсистема выключена чекбоксом, держим серый кружок
            Brush = _isChecked ? (Brush)brush : Brushes.Gray;

            foreach (var child in Children)
                child.ApplyMode(mode);
        }

        public void SetGray()
        {
            // не трогаем IsChecked/IsVisible, просто сбрасываем цвета и подсветку
            Brush = Brushes.Gray;
            foreach (var child in Children)
                child.SetGray();
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    /// <summary>
    /// Дочерний элемент легенды: отдельная линия.
    /// </summary>
    public class SubsystemLegendItem : INotifyPropertyChanged
    {
        private readonly SystemsHighlighterControlViewModel _parentVm;
        private readonly IModelViewer _viewer;

        private Brush _brush = Brushes.Gray;
        private bool _isVisible;   // чекбокс у строки

        public string Label { get; }
        public IModelElementId[] ElementIds { get; }

        public double WeldingPercent { get; }
        public double NdtPercent { get; }
        public double TestsPercent { get; }

        public Brush Brush
        {
            get => _brush;
            private set
            {
                if (Equals(_brush, value)) return;
                _brush = value;
                OnPropertyChanged(nameof(Brush));
            }
        }

        public bool IsVisible
        {
            get => _isVisible;
            set
            {
                if (_isVisible == value) return;
                _isVisible = value;
                OnPropertyChanged(nameof(IsVisible));

                if (!_isVisible)
                {
                    // выключили чекбокс — убрать подсветку и сделать кружок серым
                    Brush = Brushes.Gray;
                    _parentVm.ClearByElementIds(ElementIds);

                    // если включён режим скрытия, реально прячем геометрию
                    if (_parentVm.HideUnhighlightedMode && _parentVm.ModelViewer != null)
                    {
                        _parentVm.ModelViewer.Hide(ElementIds);
                    }
                }
                else
                {
                    // включили чекбокс — при HideUnhighlighted_mode сразу показываем
                    if (_parentVm.HideUnhighlightedMode && _parentVm.ModelViewer != null)
                    {
                        _parentVm.ModelViewer.Show(ElementIds);
                    }

                    // если режим уже выбран, сразу красим и подсвечиваем
                    var mode = _parentVm.CurrentSubsystemViewMode;
                    if (mode != SystemsHighlighterControlViewModel.SubsystemStatusViewMode.None)
                        ApplyMode(mode);
                }
            }
        }

        public SubsystemLegendItem(
            string label,
            IModelElementId[] elementIds,
            double weldingPercent,
            double ndtPercent,
            double testsPercent,
            SystemsHighlighterControlViewModel parentVm,
            IModelViewer viewer)
        {
            Label = label;
            ElementIds = elementIds ?? Array.Empty<IModelElementId>();
            WeldingPercent = weldingPercent;
            NdtPercent = ndtPercent;
            TestsPercent = testsPercent;
            _parentVm = parentVm;
            _viewer = viewer;
        }

        /// <summary>
        /// Применить текущий режим. Учитывает IsVisible: если чекбокс снят, не красим.
        /// </summary>
        public void ApplyMode(SystemsHighlighterControlViewModel.SubsystemStatusViewMode mode)
        {
            if (!_isVisible)
            {
                // элемент выключен чекбоксом — гарантированно серый и без подсветки
                Brush = Brushes.Gray;
                _parentVm.ClearByElementIds(ElementIds);
                return;
            }

            double p = 0.0;
            switch (mode)
            {
                case SystemsHighlighterControlViewModel.SubsystemStatusViewMode.Welding:
                    p = WeldingPercent;
                    break;
                case SystemsHighlighterControlViewModel.SubsystemStatusViewMode.Ndt:
                    p = NdtPercent;
                    break;
                case SystemsHighlighterControlViewModel.SubsystemStatusViewMode.Tests:
                    p = TestsPercent;
                    break;
            }

            var color = _parentVm.GetSubsystemColor(p / 100.0);
            var wpfColor = System.Windows.Media.Color.FromArgb(color.A, color.R, color.G, color.B);
            var brush = new SolidColorBrush(wpfColor);
            brush.Freeze();
            Brush = brush;

            _parentVm.HighlightByElementIds(
                ElementIds,
                ColorCircle.FromArgb(wpfColor.A, wpfColor.R, wpfColor.G, wpfColor.B));
        }

        public void SetGray()
        {
            // только визуальный сброс
            Brush = Brushes.Gray;
            _parentVm.ClearByElementIds(ElementIds);
        }

        public event PropertyChangedEventHandler PropertyChanged;
        private void OnPropertyChanged(string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }





    public class TP_LegendItem : INotifyPropertyChanged
    {
        private readonly SystemsHighlighterControlViewModel _parentVm;
        private bool _isVisible = true;
        private Brush _nodeBrush;
        private IModelViewer _viewer;
        private bool HideUnhighlighted = false;

        public TP_LegendItem(
            string label,
            Brush originalBrush,
            IModelElementId[] elementIds,
            SystemsHighlighterControlViewModel parentVm,
            IModelViewer viewer,
            bool HU)
        {
            Label = label;
            OriginalBrush = originalBrush;
            _nodeBrush = originalBrush;
            ElementIds = elementIds;
            _parentVm = parentVm;
            _viewer = viewer;
            HideUnhighlighted = HU;
        }

        public string Label { get; }
        public Brush OriginalBrush { get; }
        public IModelElementId[] ElementIds { get; }

        public Brush Brush
        {
            get => _nodeBrush;
            private set
            {
                if (_nodeBrush == value) return;
                _nodeBrush = value;
                PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(Brush)));
            }
        }

        public bool IsVisible
        {
            get => _isVisible;
            set
            {
                if (_isVisible == value) return;
                _isVisible = value;
                OnPropertyChanged();

                if (_isVisible)
                {
                    ApplyHighlight(OriginalBrush);
                    if (HideUnhighlighted) _viewer.Show(ElementIds.GroupBy(e => e.ElementId)
                                                                .Select(g => g.First())
                                                                .ToArray());
                }
                else
                {
                    ClearHighlight();
                    if (HideUnhighlighted) _viewer.Hide(ElementIds.GroupBy(e => e.ElementId)
                                                            .Select(g => g.First())
                                                            .ToArray());
                }
                    

                _parentVm.TP_RecalculatePhysicalSums();
            }
        }

        public void HideU(bool visible)
        {
            HideUnhighlighted = visible;
        }

        private void ApplyHighlight(Brush brush)
        {
            // меняем цвет кружка в легенде
            Brush = brush;
            // прокидываем в 3D
            var c = (SolidColorBrush)brush;
            _parentVm.HighlightByElementIds(ElementIds,
                ColorCircle.FromArgb(c.Color.A, c.Color.R, c.Color.G, c.Color.B));
        }

        private void ClearHighlight()
        {
            // серый в легенде
            Brush = Brushes.Gray;
            // очищаем в 3D (либо перекрашиваем серым, либо снимаем подсветку)
            _parentVm.ClearByElementIds(ElementIds);
        }

        private void OnPropertyChanged([CallerMemberName] string p = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(p));

        public event PropertyChangedEventHandler PropertyChanged;
    }

    // ====== ДОЧЕРНИЙ УЗЕЛ: СЕКЦИЯ ======
    public class PP_LegendItem : INotifyPropertyChanged
    {
        private readonly SystemsHighlighterControlViewModel _parentVm;
        private bool _isVisible = true;
        private Brush _nodeBrush;
        private IModelViewer _viewer;
        private bool HideUnhighlighted;

        // --- НОВОЕ: ссылка на родителя для наследования цвета ---
        private PP_PassportLegendNode _parentNode;

        public PP_LegendItem(
            string label,
            Brush originalBrush,
            IModelElementId[] elementIds,
            SystemsHighlighterControlViewModel parentVm,
            IModelViewer viewer,
            bool hideUnhighlighted)
        {
            Label = label;
            OriginalBrush = originalBrush;    // свой "личный" цвет секции
            _nodeBrush = originalBrush;
            ElementIds = elementIds;
            _parentVm = parentVm;
            _viewer = viewer;
            HideUnhighlighted = hideUnhighlighted;
        }

        public void AttachParent(PP_PassportLegendNode parent)
        {
            _parentNode = parent;
            // реагируем на смену IsChecked/Brush родителя
            _parentNode.PropertyChanged += (s, e) =>
            {
                if (e.PropertyName == "IsChecked" || e.PropertyName == "Brush")
                    OnParentStateChanged();
            };
            // при первичном подключении тоже актуализируемся
            OnParentStateChanged();
        }

        public string Label { get; private set; }
        public Brush OriginalBrush { get; private set; }
        public IModelElementId[] ElementIds { get; private set; }

        public Brush Brush
        {
            get { return _nodeBrush; }
            private set
            {
                if (_nodeBrush == value) return;
                _nodeBrush = value;
                var h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs(nameof(Brush)));
            }
        }

        public bool IsVisible
        {
            get { return _isVisible; }
            set
            {
                if (_isVisible == value) return;
                _isVisible = value;
                OnPropertyChanged();

                if (_isVisible)
                {
                    // ВКЛ: подсветить цветом, учитывая наследование от родителя
                    ApplyHighlight(GetEffectiveBrush());
                    if (HideUnhighlighted) _viewer.Show(ElementIds.GroupBy(e => e.ElementId).Select(g => g.First()).ToArray());
                }
                else
                {
                    // ВЫКЛ: сереем и снимаем подсветку
                    ClearHighlight();
                    if (HideUnhighlighted) _viewer.Hide(ElementIds.GroupBy(e => e.ElementId).Select(g => g.First()).ToArray());
                }

                _parentVm.PP_RecalculatePhysicalSumms();
            }
        }

        public void HideU(bool visible)
        {
            HideUnhighlighted = visible;
        }

        // --- НОВОЕ: вычисление эффективного цвета (наследуем от родителя, если он включен) ---
        private Brush GetEffectiveBrush()
        {
            if (_parentNode != null && _parentNode.IsChecked)
                return _parentNode.Brush; // цвет родителя (не серый, когда включен)
            return OriginalBrush;          // собственный цвет секции
        }

        // --- НОВОЕ: дергается, когда у родителя сменился чекбокс/цвет ---
        public void OnParentStateChanged()
        {
            if (IsVisible)
                ApplyHighlight(GetEffectiveBrush()); // перекрасим легенду и 3D в актуальный цвет
            else
                Brush = Brushes.Gray;                // выключенные секции — серые
        }

        private void ApplyHighlight(Brush brush)
        {
            // цвет кружка в легенде
            Brush = brush;

            // прокинем тот же цвет в 3D
            var c = (SolidColorBrush)brush;
            _parentVm.HighlightByElementIds(
                ElementIds,
                ColorCircle.FromArgb(c.Color.A, c.Color.R, c.Color.G, c.Color.B));
        }

        private void ClearHighlight()
        {
            // серый в легенде
            Brush = Brushes.Gray;
            // снять подсветку в 3D
            _parentVm.ClearByElementIds(ElementIds);
        }

        private void OnPropertyChanged([CallerMemberName] string p = null)
        {
            var h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs(p));
        }

        public event PropertyChangedEventHandler PropertyChanged;
    }


    // ====== КЛАСС РОДИТЕЛЬСКОГО УЗЛА: ПАСПОРТ (tri-state чекбокс) ======
    public class PP_PassportLegendNode : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler PropertyChanged;
        private void Raise(string name) { var h = PropertyChanged; if (h != null) h(this, new PropertyChangedEventArgs(name)); }

        public string Label { get; private set; }
        public ObservableCollection<PP_LegendItem> Children { get; private set; }

        // Цвет родителя в "включенном" состоянии (его фирменный цвет)
        public Brush OriginalBrush { get; private set; }

        // Текущий цвет кружка родителя (серый, если выключен)
        private Brush _brush;
        public Brush Brush
        {
            get { return _brush; }
            private set { if (!Equals(_brush, value)) { _brush = value; Raise(nameof(Brush)); } }
        }

        public PP_PassportLegendNode(string passportLabel, Brush brush)
        {
            Label = passportLabel;
            Children = new ObservableCollection<PP_LegendItem>();
            OriginalBrush = brush;
            _brush = Brushes.Gray;       // по умолчанию серый
            _isChecked = false;          // по умолчанию выключен
        }

        private bool _isExpanded = true;
        public bool IsExpanded
        {
            get { return _isExpanded; }
            set { if (_isExpanded != value) { _isExpanded = value; Raise(nameof(IsExpanded)); } }
        }

        // Только вкл/выкл
        private bool _isChecked;
        public bool IsChecked
        {
            get { return _isChecked; }
            set
            {
                if (_isChecked == value) return;
                _isChecked = value;
                Brush = _isChecked ? OriginalBrush : Brushes.Gray;   // <-- динамический цвет кружка
                Raise(nameof(IsChecked));

                // клик по родителю включает/выключает всех детей
                foreach (var ch in Children)
                    ch.IsVisible = value;

                // на всякий случай обновим цвет/подсветку видимых детей
                NotifyChildrenParentState();
                UpdateCheckedFromChildren(); // синхронизация, если что-то поменяется в процессе
            }
        }

        public void HookChildren()
        {
            foreach (var ch in Children)
            {
                ch.AttachParent(this); // <-- сообщаем детям, кто их родитель
                ch.PropertyChanged += (s, e) =>
                {
                    if (e.PropertyName == nameof(PP_LegendItem.IsVisible))
                        UpdateCheckedFromChildren();
                };
            }
            UpdateCheckedFromChildren();
        }

        public void UpdateCheckedFromChildren()
        {
            if (Children.Count == 0)
            {
                if (_isChecked != false) { _isChecked = false; Brush = Brushes.Gray; Raise(nameof(IsChecked)); }
                return;
            }

            int on = Children.Count(c => c.IsVisible);
            bool newState = (on == Children.Count); // родитель "вкл" только если включены все секции

            if (_isChecked != newState)
            {
                _isChecked = newState;
                Brush = _isChecked ? OriginalBrush : Brushes.Gray; // кружок родителя
                Raise(nameof(IsChecked));
                NotifyChildrenParentState(); // у видимых детей надо обновить цвет (наследуем/возвращаем)
            }
        }

        private void NotifyChildrenParentState()
        {
            foreach (var ch in Children)
                ch.OnParentStateChanged();
        }
    }



    public class SystemNodeViewModel : INotifyPropertyChanged
    {
        private readonly SystemsHighlighterControlViewModel _parentVm;
        public IModelElementId[] _elementIds;
        private bool _isHighlighted;
        private Brush _nodeBrush = new SolidColorBrush(ColorCircle.FromArgb(255, 200, 200, 200));
        private IModelViewer _viewer;

        // Правила жёсткой привязки по подстроке -> цвет
        private static readonly (string Pattern, ColorCircle Color)[] FixedColorRules = new[]
        {
            // Вода 
            ("DW", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("PW", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("SRW", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("TW", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("ADS", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("BFW", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("STW", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("CWI", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("CWII", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("CWRI", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("CWRII", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("CWIII", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("CWRIII", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("RW", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("W18", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("W21", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("W7", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("W8", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("WD", ColorCircle.FromArgb(255, 11, 83, 61)),
            ("WF", ColorCircle.FromArgb(255, 11, 83, 61)),

            // Пар
            ("HS", ColorCircle.FromArgb(255, 231, 47, 37)),
            ("HSS", ColorCircle.FromArgb(255, 231, 47, 37)),
            ("HС", ColorCircle.FromArgb(255, 231, 47, 37)),
            ("LS", ColorCircle.FromArgb(255, 231, 47, 37)),
            ("LSS", ColorCircle.FromArgb(255, 231, 47, 37)),
            ("LC", ColorCircle.FromArgb(255, 231, 47, 37)),
            ("MS", ColorCircle.FromArgb(255, 231, 47, 37)),
            ("MSS", ColorCircle.FromArgb(255, 231, 47, 37)),
            ("MС", ColorCircle.FromArgb(255, 231, 47, 37)),

            // Воздух
            ("VA", ColorCircle.FromArgb(255, 55, 99, 175)),
            ("IA", ColorCircle.FromArgb(255, 55, 99, 175)),
            ("ВА", ColorCircle.FromArgb(255, 55, 99, 175)),
            ("РА", ColorCircle.FromArgb(255, 55, 99, 175)),


            // Газы
            ("AG", ColorCircle.FromArgb(255, 249, 239, 71)),
            ("FG", ColorCircle.FromArgb(255, 249, 239, 71)),
            ("FGV", ColorCircle.FromArgb(255, 249, 239, 71)),
            ("NG", ColorCircle.FromArgb(255, 249, 239, 71)),
            ("IF", ColorCircle.FromArgb(255, 249, 239, 71)),
            ("HI", ColorCircle.FromArgb(255, 249, 239, 71)),
            ("LI", ColorCircle.FromArgb(255, 249, 239, 71)),
            ("FA", ColorCircle.FromArgb(255, 249, 239, 71)),
            ("FL", ColorCircle.FromArgb(255, 249, 239, 71)),
            ("FН", ColorCircle.FromArgb(255, 249, 239, 71)),
            ("Н", ColorCircle.FromArgb(255, 249, 239, 71)),

            // Жидкости
            ("FО", ColorCircle.FromArgb(255, 137, 105, 65)),
            ("CL", ColorCircle.FromArgb(255, 137, 105, 65)),
            ("SO", ColorCircle.FromArgb(255, 137, 105, 65)),
            ("ОС", ColorCircle.FromArgb(255, 137, 105, 65)),
            ("IO", ColorCircle.FromArgb(255, 137, 105, 65)),
            ("СО", ColorCircle.FromArgb(255, 137, 105, 65)),
            ("W19", ColorCircle.FromArgb(255, 137, 105, 65)),
            ("DEA", ColorCircle.FromArgb(255, 137, 105, 65)),
            ("GL", ColorCircle.FromArgb(255, 137, 105, 65)),
            ("MDEA", ColorCircle.FromArgb(255, 137, 105, 65)),
            ("MEA", ColorCircle.FromArgb(255, 137, 105, 65)),
            ("D54", ColorCircle.FromArgb(255, 137, 105, 65)),

            //Кислоты 
            ("SA", ColorCircle.FromArgb(255, 241, 187, 23)),

            //Щелочи
            ("ZS", ColorCircle.FromArgb(255, 137, 59, 123)),

        };

        public SystemNodeViewModel(
            string name,
            IModelElementId[] elementIds,
            SystemsHighlighterControlViewModel parentVm)
        {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            _elementIds = elementIds ?? throw new ArgumentNullException(nameof(elementIds));
            _parentVm = parentVm ?? throw new ArgumentNullException(nameof(parentVm));
            Children = new ObservableCollection<SystemNodeViewModel>();

            // По умолчанию нули, чтобы биндинг не был пустым
            SetTotals(new PhysicalTotals(0, 0, 0, 0));
        }

        public string Name { get; }
        public ObservableCollection<SystemNodeViewModel> Children { get; }

        public Brush NodeBrush
        {
            get => _nodeBrush;
            private set
            {
                if (_nodeBrush == value) return;
                _nodeBrush = value;
                OnPropertyChanged();
            }
        }

        /// <summary>
        /// Применяет подсветку заданным цветом к этому узлу и оповещает ViewModel
        /// о добавлении/удалении для пересчёта физических сумм.
        /// </summary>
        public void ApplyHighlight(bool isOn, ColorCircle brushColor)
        {
            _isHighlighted = isOn;
            OnPropertyChanged(nameof(IsHighlighted));

            NodeBrush = new SolidColorBrush(brushColor);

            if (isOn)
                _parentVm.HighlightByElementIds(_elementIds, brushColor);
            else
                _parentVm.ClearByElementIds(_elementIds);

            _parentVm.RecalculatePhysicalSums();
        }

        /// <summary>
        /// Двусторонний биндинг для чекбокса.
        /// При изменении состояния запускает подсветку для себя и детей.
        /// </summary>
        public bool IsHighlighted
        {
            get => _isHighlighted;
            set
            {
                if (_isHighlighted == value) return;

                // Определяем цвет: сначала по правилам FixedColorRules, иначе у parentVm запрашиваем следующий уникальный
                ColorCircle color;
                if (value)
                {
                    var fixedRule = Array.Find(FixedColorRules, r =>
                        Name.IndexOf(r.Pattern, StringComparison.OrdinalIgnoreCase) >= 0);
                    if (!fixedRule.Equals(default))
                        color = fixedRule.Color;
                    else
                        color = _parentVm.GetNextUniqueColor();
                }
                else
                {
                    color = ColorCircle.FromArgb(255, 200, 200, 200);
                }

                ApplyHighlightRecursive(this, value, color);

                if (_viewer == null)
                {
                    _viewer = _parentVm.GetModelViewer();
                }

                if (value)
                {
                    ShowRecursive(this);
                }
                else if (_parentVm.HideUnhighlighted_mode)
                {
                    HideRecursive(this);
                }
            }
        }

        // Рекурсивная окраска
        private void ApplyHighlightRecursive(SystemNodeViewModel node, bool highlight, ColorCircle color)
        {
            node.ApplyHighlight(highlight, color);
            foreach (var child in node.Children)
                ApplyHighlightRecursive(child, highlight, color);
        }

        // Рекурсивный показ
        private void ShowRecursive(SystemNodeViewModel node)
        {
            if (node.Children.Count == 0)
                _viewer.Show(node._elementIds);

            foreach (var child in node.Children)
                ShowRecursive(child);
        }

        // Рекурсивное скрытие
        private void HideRecursive(SystemNodeViewModel node)
        {
            if (node.Children.Count == 0)
                _viewer.Hide(node._elementIds);

            foreach (var child in node.Children)
                HideRecursive(child);
        }

        public void HideUnhighlighted(IModelViewer mv)
        {
            mv.Hide(_elementIds);
        }

        public void ReturnVisibility(IModelViewer mv)
        {
            mv.Show(_elementIds);
        }


        public PhysicalTotals Totals { get; private set; }

        // Для биндинга/отображения в дереве:
        private string _weight, _length, _volume, _diaInch;
        public string Weight { get => _weight; private set { _weight = value; OnPropertyChanged(); OnPropertyChanged(nameof(Title)); } }
        public string Length { get => _length; private set { _length = value; OnPropertyChanged(); OnPropertyChanged(nameof(Title)); } }
        public string Volume { get => _volume; private set { _volume = value; OnPropertyChanged(); OnPropertyChanged(nameof(Title)); } }
        public string DiaInch { get => _diaInch; private set { _diaInch = value; OnPropertyChanged(); OnPropertyChanged(nameof(Title)); } }

        // Итоговый текст для узла (если хотите показывать суммы рядом с именем)
        public string Title =>
            $"{Name}  |  T={Weight}  L={Length}  V={Volume}  DI={DiaInch}";

        public void SetTotals(PhysicalTotals totals)
        {
            Totals = totals;
            Weight = Fmt(totals.Weight);
            Length = Fmt(totals.Length);
            Volume = Fmt(totals.Volume);
            DiaInch = Fmt(totals.DiaInch);
        }

        public void RecomputeFromChildren()
        {
            if (Children.Count == 0) return;
            PhysicalTotals sum = default;
            foreach (var c in Children) sum += c.Totals;
            SetTotals(sum);
        }



        #region INotifyPropertyChanged
        public event PropertyChangedEventHandler PropertyChanged;
        protected void OnPropertyChanged([CallerMemberName] string propName = null) =>
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propName));
        #endregion
    }

    public class DateThresholdBucket
    {
        public DateTime Threshold { get; }
        public IModelElementId[] ElementIds { get; }

        public DateThresholdBucket(DateTime threshold, IModelElementId[] elementIds)
        {
            Threshold = threshold;
            ElementIds = elementIds;
        }
    }



    public class PdfReportBuilder
    {
        private readonly string _title;
        private readonly DateTime _createdAt;
        private readonly string[] _headers;
        private readonly List<string[]> _rows;
        private readonly byte[] _logoBytes;

        // нормализованные данные
        private string[] _normHeaders;
        private List<string[]> _normRows;
        private bool[] _numericCols;
        private int _colCount;
        private int _textColsCount;

        public PdfReportBuilder(string title, DateTime createdAt, string[] headers, List<string[]> rows)
        {
            _title = title ?? "";
            _createdAt = createdAt;

            _headers = headers != null ? headers : new string[0];
            _rows = rows != null ? rows : new List<string[]>();

            // подгружаем логотип (может быть null — это ок)
            _logoBytes = IconLoader.GetIcon("logo.png");
        }

        public void GeneratePdf(string filePath)
        {
            QuestPDF.Settings.License = LicenseType.Community;

            PrepareData();

            Document.Create(container =>
            {
                container.Page(page =>
                {
                    page.Size(PageSizes.A4.Landscape());
                    page.Margin(20);
                    page.DefaultTextStyle(TextStyle.Default.FontSize(9).FontFamily("Arial"));

                    page.Header().Element(c => BuildHeader(c));
                    page.Content().Element(c => BuildTable(c));

                    page.Footer().Height(36).PaddingTop(4).Row(row =>
                    {
                        row.RelativeItem().AlignCenter().AlignMiddle().Text(t =>
                        {
                            t.Span("Стр. ");
                            t.CurrentPageNumber();
                            t.Span(" / ");
                            t.TotalPages();
                        });

                        if (_logoBytes != null)
                        {
                            row.ConstantItem(72).AlignRight().AlignBottom().Element(c =>
                            {
                                c.Width(64).Height(30).Image(_logoBytes).FitWidth();
                            });
                        }
                    });
                });
            }).GeneratePdf(filePath);
        }

        // ---------- Подготовка данных ----------

        private void PrepareData()
        {
            _colCount = Math.Max(_headers.Length, MaxRowLength(_rows));
            if (_colCount == 0)
                throw new InvalidOperationException("Нет данных для PDF: пустые заголовки и строки.");

            _normHeaders = NormalizeArray(_headers, _colCount);
            _normRows = new List<string[]>(_rows.Count);
            for (int i = 0; i < _rows.Count; i++)
            {
                string[] r = _rows[i] ?? new string[0];
                _normRows.Add(NormalizeArray(r, _colCount));
            }

            _numericCols = DetectNumericColumns(_normRows, _colCount);
            _textColsCount = 0;
            for (int i = 0; i < _colCount; i++)
                if (!_numericCols[i]) _textColsCount++;
        }

        private int MaxRowLength(List<string[]> list)
        {
            int max = 0;
            if (list == null) return 0;
            for (int i = 0; i < list.Count; i++)
            {
                string[] arr = list[i];
                int len = arr != null ? arr.Length : 0;
                if (len > max) max = len;
            }
            return max;
        }

        private string[] NormalizeArray(string[] source, int length)
        {
            string[] result = new string[length];
            int copy = source != null ? Math.Min(source.Length, length) : 0;
            if (copy > 0) Array.Copy(source, result, copy);
            for (int i = 0; i < length; i++)
                if (result[i] == null) result[i] = string.Empty;
            return result;
        }

        private bool[] DetectNumericColumns(List<string[]> rows, int colCount)
        {
            bool[] numeric = new bool[colCount];
            for (int c = 0; c < colCount; c++)
            {
                int total = 0;
                int ok = 0;

                for (int r = 0; r < rows.Count; r++)
                {
                    string[] row = rows[r];
                    if (row == null || c >= row.Length) continue;
                    string s = (row[c] ?? "").Trim();
                    if (s.Length == 0) continue;

                    total++;
                    if (LooksLikeNumber(s)) ok++;
                }

                // считаем колонку числовой, если >= 90% непустых значений парсятся как число
                numeric[c] = (total > 0) && (ok * 10 >= total * 9);
            }
            return numeric;
        }

        private bool LooksLikeNumber(string s)
        {
            double d;
            if (double.TryParse(s, NumberStyles.Any, CultureInfo.InvariantCulture, out d))
                return true;
            if (double.TryParse(s, NumberStyles.Any, new CultureInfo("ru-RU"), out d))
                return true;
            return false;
        }

        // ---------- Хедер ----------

        private void BuildHeader(IContainer container)
        {
            container.Column(col =>
            {
                col.Spacing(0);
                col.Item().Row(row =>
                {
                    row.RelativeItem()
                       .Text(_title)
                       .Style(TextStyle.Default.FontSize(13).SemiBold());

                    row.ConstantItem(200)
                       .AlignRight()
                       .Text(string.Format("Создано: {0:dd.MM.yyyy HH:mm}", _createdAt));
                });
            });
        }

        // ---------- Таблица ----------

        private IContainer HeaderCellBase(IContainer c)
        {
            return c.PaddingVertical(0).PaddingHorizontal(0)
                    .Background(QuestPDF.Helpers.Colors.Grey.Lighten3)
                    .Border(0.3f).BorderColor(QuestPDF.Helpers.Colors.Grey.Lighten2)
                    .AlignMiddle().AlignCenter();
        }

        private IContainer BodyCellBase(IContainer c, bool isTotal)
        {
            return c.PaddingVertical(0).PaddingHorizontal(0)
                    .Background(isTotal ? QuestPDF.Helpers.Colors.Grey.Lighten3 : QuestPDF.Helpers.Colors.White)
                    .Border(0.25f).BorderColor(QuestPDF.Helpers.Colors.Grey.Lighten4)
                    .MinHeight(12)
                    .AlignMiddle() 
                    .AlignCenter(); 
        }

        private void BuildTable(IContainer container)
        {
            int colCount = _colCount;

            container.Table(table =>
            {
                // колонки
                table.ColumnsDefinition(cols =>
                {
                    for (int i = 0; i < colCount; i++)
                    {
                        if (_numericCols[i]) cols.ConstantColumn(90);
                        else cols.RelativeColumn(3);
                    }
                });

                // шапка — один контент на ячейку
                table.Header(header =>
                {
                    for (int i = 0; i < colCount; i++)
                    {
                        header.Cell()
                              .Element(HeaderCellBase)
                              .Text(_normHeaders[i]);
                    }
                });

                // тело
                for (int r = 0; r < _normRows.Count; r++)
                {
                    string[] row = _normRows[r];
                    bool isTotal = row.Length > 0 &&
                                   string.Equals((row[0] ?? "").Trim(), "ИТОГО:", StringComparison.OrdinalIgnoreCase);

                    if (isTotal && _textColsCount > 1)
                    {
                        int firstNumericIndex = FirstNumericIndex();
                        if (firstNumericIndex < 0) firstNumericIndex = colCount;

                        table.Cell()
                             .ColumnSpan((uint)_textColsCount)
                             .Element(delegate (IContainer e) { return BodyCellBase(e, true); })
                             .Text(row[0].Length > 0 ? row[0] : "ИТОГО:");

                        for (int c = firstNumericIndex; c < colCount; c++)
                        {
                            table.Cell()
                                 .Element(delegate (IContainer e)
                                 {
                                     IContainer t = BodyCellBase(e, true);
                                     if (_numericCols[c]) t = t.AlignRight();
                                     return t;
                                 })
                                 .Text(row[c]);
                        }
                    }
                    else
                    {
                        for (int c = 0; c < colCount; c++)
                        {
                            string cellText = row[c];

                            table.Cell()
                                 .Element(delegate (IContainer e)
                                 {
                                     IContainer t = BodyCellBase(e, isTotal);
                                     //if (_numericCols[c]) t = t.AlignRight();
                                     return t;
                                 })
                                 .Text(cellText);
                        }
                    }
                }
            });
        }

        private int FirstNumericIndex()
        {
            for (int i = 0; i < _numericCols.Length; i++)
                if (_numericCols[i]) return i;
            return -1;
        }
    }

    public static class SubsystemsPdfReportHelper
    {
        /// <summary>
        /// Генерирует PDF-отчёт по подсистемам на основе mapper'а и статусов по линиям.
        /// </summary>
        public static void GenerateSubsystemsPdfReport(
            SubsystemsStatusMapper mapper,
            Dictionary<string, SubsystemsStatusReader.LineStatus> lineStatuses,
            string filePath)
        {
            if (mapper == null)
                throw new ArgumentNullException(nameof(mapper));
            if (lineStatuses == null)
                throw new ArgumentNullException(nameof(lineStatuses));
            if (string.IsNullOrWhiteSpace(filePath))
                throw new ArgumentException("Путь к PDF не задан.", nameof(filePath));

            // 1. Агрегируем данные по подсистемам
            var subsystemsData = mapper.BuildSubsystemStatuses(lineStatuses);

            // 2. Готовим заголовки таблицы
            var headers = new[]
            {
                "Система",
                "Подсистема",
                "Приоритеты",
                "Тест-пакеты",
                "Средняя готовность по сварке, %",
                "Средняя готовность по НК, %",
                "Испытания завершены, % строк с \"Да\"",
                "Кол-во элементов"
            };

            // 3. Формируем строки
            var rows = new List<string[]>();

            foreach (var sysKv in subsystemsData.OrderBy(k => k.Key))
            {
                var systemName = sysKv.Key;
                var subsDict = sysKv.Value;

                foreach (var subsKv in subsDict.OrderBy(k => k.Key))
                {
                    var s = subsKv.Value;

                    string avgWelding = s.AvgWeldingPercent.ToString("0.0", CultureInfo.InvariantCulture);
                    string avgNdt = s.AvgNdtPercent.ToString("0.0", CultureInfo.InvariantCulture);
                    string testsCompleted = s.TestsCompletedPercent.ToString("0.0", CultureInfo.InvariantCulture);
                    string elementsCount = (s.Elements != null ? s.Elements.Count : 0).ToString(CultureInfo.InvariantCulture);

                    rows.Add(new[]
                    {
                        s.SystemName ?? systemName ?? string.Empty,
                        s.SubsystemName ?? subsKv.Key ?? string.Empty,
                        s.PrioritiesText ?? string.Empty,
                        s.TestPackagesText ?? string.Empty,
                        avgWelding,
                        avgNdt,
                        testsCompleted,
                        elementsCount
                    });
                }
            }

            // 4. Создаём и генерим PDF
            var title = "Статус подсистем по тест-пакетам";
            var createdAt = DateTime.Now;

            var builder = new PdfReportBuilder(title, createdAt, headers, rows);
            builder.GeneratePdf(filePath);
        }
    }



}