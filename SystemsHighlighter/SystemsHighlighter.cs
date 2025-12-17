using System;
using System.ComponentModel.Composition;
using System.Linq;
using System.Windows;
using Ascon.Pilot.SDK;                  // IDataPlugin
using Ascon.Pilot.Bim.SDK;
using Ascon.Pilot.Bim.SDK.ModelTab;     // ITabsManager, IModelTab
using Ascon.Pilot.Bim.SDK.Model;
using System.Collections.Generic;
using Ascon.Pilot.Bim.SDK.ModelViewer;  // IModelManager
using System.Windows.Threading;         // Dispatcher

namespace SystemsHighlighter
{
    [Export(typeof(IDataPlugin))]
    public class SystemsHighlighter : IDataPlugin
    {
        public const string TITLE = nameof(SystemsHighlighter);
        private readonly byte[] _icon;

        private readonly ITabsManager _tabsManager;
        private readonly IModelManager _modelManager;

        public string _filePath;
        public List<IType> Types { get; }

        private readonly Dictionary<Guid, SystemsHighlighterControlViewModel> _tabViewModels = new Dictionary<Guid, SystemsHighlighterControlViewModel>();
        private SystemsHighlighterControlViewModel _viewModel;

        private readonly List<Ascon.Pilot.SDK.IDataObject> _elements;
        private readonly IObjectsRepository _objectsRepository;

        // UI
        private readonly Dispatcher _dispatcher;

        // Подписки
        private IDisposable _rootSub;
        private IDisposable _childrenSub;

        // Идентификаторы
        private readonly Guid _rootId = new Guid("d8ea8564-2e67-4f5f-98bc-a0f3f5ed54b8"/*"c36aea43-ce60-4de2-bf8a-bf1e8ea3713b"*/);

        // Синхронизация
        private readonly object _guard = new object();

        // Буфер актуальных данных
        private volatile bool _dataReady;
        private List<Ascon.Pilot.SDK.IDataObject> _lastItems = new List<Ascon.Pilot.SDK.IDataObject>();

        // Доставить данные во все живые VM (с учётом фильтра)
        private void PushDataToAllVMs()
        {
            List<Ascon.Pilot.SDK.IDataObject> snapshot;
            lock (_guard)
                snapshot = _lastItems?.ToList() ?? new List<Ascon.Pilot.SDK.IDataObject>();

            // Централизованный фильтр по "Piping" (без учёта регистра)
            var payload = snapshot.Where(e => e.DisplayName?.Contains("Piping") == true)
                                  .ToList();

            // Если данных ещё реально нет и мы не в состоянии "готово" — молчим
            if (payload.Count == 0 && !_dataReady)
                return;

            foreach (var vm in _tabViewModels.Values.Where(v => v != null))
                vm.get_dataobjects(payload);
        }

        [ImportingConstructor]
        public SystemsHighlighter(IPilotServiceProvider pilotServiceProvider, IObjectsRepository objectsRepository)
        {
            _tabsManager = pilotServiceProvider.GetServices<ITabsManager>().First();
            _modelManager = pilotServiceProvider.GetServices<IModelManager>().Last();
            _objectsRepository = objectsRepository;

            _icon = IconLoader.GetIcon("IconPSH.svg");

            var objType = objectsRepository.GetType("project");
            Types = new List<IType>() { objType };

            _filePath = objectsRepository.GetStoragePath() + @"Рабочая папка\";
            objectsRepository.Mount(Guid.Parse("b8a63a39-03ce-4b13-bed8-116e9b54b0da"));

            _elements = new List<Ascon.Pilot.SDK.IDataObject>();

            _dispatcher = Application.Current?.Dispatcher
                ?? throw new InvalidOperationException("UI Dispatcher недоступен. Инициализируйте SystemsHighlighter на UI-потоке.");

            // Сайдбар сначала инициализируем для уже открытых вкладок,
            // затем подписываемся на жизненный цикл вкладок/моделей
            var openedTabs = _tabsManager.GetTabs().OfType<IModelTab>().ToList();
            InitOpenedTabsSidebar(openedTabs);
            SubscribeTabLifecycle();

            // Запускаем двухфазную загрузку: root -> children
            StartRootAndChildrenFlow();
        }

        // === Старт: подписка на корень (НЕ одноразовая) ===
        private void StartRootAndChildrenFlow()
        {
            var rootObs = new RootObserver(this);
            _rootSub = _objectsRepository.SubscribeObjects(new[] { _rootId }).Subscribe(rootObs);
            rootObs.Subscription = _rootSub;
        }

        // Вызывается RootObserver при новых снапшотах корня
        internal void OnRootReceived(Ascon.Pilot.SDK.IDataObject root)
        {
            // Собираем уникальные Guid детей
            var childIds = root?.Children != null
                ? new HashSet<Guid>(root.Children)
                : new HashSet<Guid>();

            lock (_guard)
            {
                _elements.Clear();
            }

            if (childIds.Count == 0)
            {
                // Нет детей — считаем данные готовыми, но пустыми; очищаем VM
                lock (_guard)
                {
                    _lastItems = new List<Ascon.Pilot.SDK.IDataObject>();
                    _dataReady = true;
                }

                _dispatcher.BeginInvoke(new Action(() =>
                {
                    PushDataToAllVMs(); // разослать пустой набор, чтобы UI очистился
                }), DispatcherPriority.Background);

                return;
            }

            // Подписываемся на детей: ждём, пока придут все хотя бы по разу
            var childrenObs = new ChildrenObserver(this, childIds);
            _childrenSub?.Dispose(); // перестраховка от висящих предыдущих подписок
            _childrenSub = _objectsRepository.SubscribeObjects(childIds).Subscribe(childrenObs);
            childrenObs.Subscription = _childrenSub;
        }

        // Вызывается ChildrenObserver, когда собран полный набор детей
        internal void OnChildrenComplete(IReadOnlyCollection<Ascon.Pilot.SDK.IDataObject> items)
        {
            _dispatcher.BeginInvoke(new Action(() =>
            {
                lock (_guard)
                {
                    _elements.Clear();
                    foreach (var it in items) _elements.Add(it);

                    // Обновляем буфер и помечаем как готовый
                    _lastItems = _elements.ToList();
                    _dataReady = true;
                }

                // Разослать всем живым VM
                PushDataToAllVMs();
            }), DispatcherPriority.Background);
        }

        // ===== Внутренние наблюдатели =====

        private sealed class RootObserver : IObserver<Ascon.Pilot.SDK.IDataObject>
        {
            private readonly SystemsHighlighter _owner;
            private bool _gotRootOnce;

            internal IDisposable Subscription { get; set; }

            public RootObserver(SystemsHighlighter owner) => _owner = owner;

            public void OnNext(Ascon.Pilot.SDK.IDataObject value)
            {
                if (value == null) return;

                // Первую волну детей запускаем один раз, но подписку НЕ рвём,
                // чтобы ловить последующие изменения структуры.
                if (!_gotRootOnce)
                {
                    _gotRootOnce = true;
                    _owner.OnRootReceived(value);
                }
                else
                {
                    // При каждом новом снапшоте корня перезапускаем сбор детей
                    // (можно добавить сравнение наборов для оптимизации)
                    _owner.OnRootReceived(value);
                }
            }

            public void OnError(Exception error)
            {
                System.Diagnostics.Debug.WriteLine(error);
            }

            public void OnCompleted()
            {
                // Поставщик изменений обычно не завершает поток. Игнорируем.
            }
        }

        private sealed class ChildrenObserver : IObserver<Ascon.Pilot.SDK.IDataObject>
        {
            private readonly SystemsHighlighter _owner;
            private readonly HashSet<Guid> _pending;
            private readonly Dictionary<Guid, Ascon.Pilot.SDK.IDataObject> _byId;

            internal IDisposable Subscription { get; set; }

            public ChildrenObserver(SystemsHighlighter owner, HashSet<Guid> expectedIds)
            {
                _owner = owner;
                _pending = new HashSet<Guid>(expectedIds);
                _byId = new Dictionary<Guid, Ascon.Pilot.SDK.IDataObject>(_pending.Count);
            }

            public void OnNext(Ascon.Pilot.SDK.IDataObject value)
            {
                if (value == null) return;

                _byId[value.Id] = value;
                _pending.Remove(value.Id);

                if (_pending.Count == 0)
                {
                    // Первый полный набор получили — отписываемся
                    Subscription?.Dispose();
                    _owner.OnChildrenComplete(_byId.Values.ToList());
                }
            }

            public void OnError(Exception error)
            {
                System.Diagnostics.Debug.WriteLine(error);
            }

            public void OnCompleted()
            {
                // Не рассчитываем на это событие.
            }
        }

        // ===== UI/таб-менеджмент =====

        private void SubscribeTabLifecycle()
        {
            _tabsManager.TabOpened += OnTabOpened;
            _modelManager.ModelLoaded += OnModelLoaded;
            _modelManager.ModelClosed += OnModelClosed;
        }

        private void InitOpenedTabsSidebar(List<IModelTab> openedTabs)
        {
            openedTabs.ForEach(AddSidebarTab);
        }

        private void AddSidebarTab(IModelTab modelTab)
        {
            // На всякий случай — UI dispatcher должен быть
            var dispatcher = Application.Current?.Dispatcher
                             ?? throw new InvalidOperationException("UI Dispatcher недоступен.");

            dispatcher.Invoke(() =>
            {
                // 1) Всё, что связано с модельными/визуальными менеджерами — получаем на UI-потоке
                var sidebarManager = modelTab?.GetSidebarManager();
                if (sidebarManager == null)
                    throw new InvalidOperationException($"GetSidebarManager() вернул null для вкладки {modelTab?.Id}");

                // 2) Иконка тоже может быть null, если ресурс не найден — подстрахуемся
                var icon = _icon ?? IconLoader.GetIcon("IconPSH.svg");

                // 3) Создаем VM для КОНКРЕТНОЙ вкладки
                var vm = new SystemsHighlighterControlViewModel();
                vm.get_filepath(_filePath);

                // 4) UI-контрол с DataContext
                var view = new Highlighter_Bar { DataContext = vm };

                // 5) Добавляем вкладку сайдбара
                var sidebarTab = sidebarManager.AddTab(0, TITLE, icon, view);
                if (sidebarTab == null)
                    throw new InvalidOperationException("AddTab вернул null.");

                vm.SidebarTab = sidebarTab;

                // 6) Видимость и начальная инициализация — только если Viewer уже есть
                var viewer = modelTab.ModelViewer;
                sidebarTab.IsVisible = viewer != null;
                if (viewer != null)
                    vm.OnModelLoaded(viewer);

                // 7) Сохраняем VM за конкретным TabId
                _tabViewModels[modelTab.Id] = vm;

                // 8) Обновляем ссылку на "последний активный"
                _viewModel = vm;

                // 9) Если данные уже готовы — отдай их сейчас (после того как VM зарегистрирована)
                if (_dataReady)
                    PushDataToAllVMs();
            });

            // 10) Чистим за собой при закрытии вкладки
            modelTab.Disposed += (_, __) => OnModelTabDisposed(modelTab.Id);
        }

        private void OnTabOpened(object sender, TabEventArgs e)
        {
            if (e.Tab is IModelTab modelTab)
                AddSidebarTab(modelTab);

            // На случай «холодного старта» без данных — мягкий пендель пайплайну
            if (!_dataReady)
                StartRootAndChildrenFlow();
        }

        private void OnModelLoaded(object sender, ModelEventArgs e)
        {
            if (_tabViewModels.TryGetValue(e.Viewer.TabId, out var vm) && vm != null)
            {
                vm.OnModelLoaded(e.Viewer);
                vm.get_filepath(_filePath);

                if (_dataReady)
                    PushDataToAllVMs();
            }
        }

        private void OnModelClosed(object sender, ModelEventArgs e)
        {
            if (_tabViewModels.TryGetValue(e.Viewer.TabId, out var vm) && vm != null)
                vm.OnModelClosed();
        }

        public void OnModelTabDisposed(Guid id)
        {
            _tabViewModels.Remove(id);
        }

        public string get_path()
        {
            return _filePath;
        }
    }
}
