using Ascon.Pilot.Bim.SDK.ModelViewer;
using System;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Threading;

namespace SystemsHighlighter
{
    public partial class Highlighter_Bar : UserControl
    {
        public Highlighter_Bar()
        {
            InitializeComponent();
            DataContextChanged += OnDataContextChanged;

            // Подписываемся на появление/исчезновение контрола
            this.Loaded += OnControlLoaded;
            this.Unloaded += OnControlUnloaded;
        }

        // 1) Регистрируем DependencyProperty
        public static readonly DependencyProperty ModelViewerProperty =
            DependencyProperty.Register(
                nameof(ModelViewer),
                typeof(IModelViewer),
                typeof(Highlighter_Bar),
                new PropertyMetadata(null));

        // 2) CLR-обёртка
        public IModelViewer ModelViewer
        {
            get => (IModelViewer)GetValue(ModelViewerProperty);
            private set => SetValue(ModelViewerProperty, value);
        }

        private SystemsHighlighterControlViewModel _vm;

        private void OnDataContextChanged(object sender, DependencyPropertyChangedEventArgs e)
        {
            if (_vm != null)
                _vm.PropertyChanged -= VmOnPropertyChanged;

            _vm = DataContext as SystemsHighlighterControlViewModel;

            if (_vm != null)
            {
                // Поддерживаем модель визуализатора
                ModelViewer = _vm.ModelViewer;
                _vm.PropertyChanged += VmOnPropertyChanged;

                // Дополнительно: пересчитаем суммы сразу при привязке
                _vm.RecalculatePhysicalSums();
            }
        }

        private void VmOnPropertyChanged(object sender, PropertyChangedEventArgs e)
        {
            switch (e.PropertyName)
            {
                case nameof(SystemsHighlighterControlViewModel.ModelViewer):
                    Dispatcher.Invoke(() => ModelViewer = _vm.ModelViewer);
                    break;

                // Слушаем флаг, которым VM может помечать, что нужно обновить физданные:
                case nameof(SystemsHighlighterControlViewModel.ShouldRecalcPhys):
                    _vm.RecalculatePhysicalSums();
                    break;
            }
        }

        /*
        // Вызывается, когда контрол появляется в визуальном дереве
        private void OnControlLoaded(object sender, RoutedEventArgs e)
        {


            // При появлении контрола — красим всё в серый
            _vm.InitPhysData();
            _vm.GrayAllElements();
            _vm.InitModelData();
            _vm.ApplyFilter(null);
            
        }*/

        private async void OnControlLoaded(object sender, RoutedEventArgs e)
        {
            await Dispatcher.InvokeAsync(() => { }, DispatcherPriority.Render); // дождаться отрисовки UI

            if (_vm == null)
                return;

            _vm.GrayAllElements();

            _vm.cmbModelPartsToLoadItems = new ObservableCollection<string>
                                        {
                                            "010", "200", "300", "310", "320", "330", "340",
                                            "410", "420", "500", "720", "730",
                                            "801", "802", "803", "804",
                                            "900", "910"
                                        };

            //_vm.IsLoading = true;

            //_vm.LoadStatus = "Пожалуйста, подождите. Идёт загрузка систем и подсистем...";
            //await _vm.InitModelDataAsync();

            //_vm.ApplyFilter(null);

            //_vm.LoadStatus = "Пожалуйста, подождите. Идёт загрузка данных о ФО систем...";
            //await _vm.InitPhysDataAsync();

            //_vm.LoadStatus = "Пожалуйста, подождите. Идёт загрузка данных о испытаниях...";
            //await _vm.InitTPDataAsync();

            //_vm.LoadStatus = "Все данные готовы к визуализации";
            //_vm.IsLoading = false;
        }


        private async void LoadModelPartAsync()
        {
            _vm.IsLoading = true;

            _vm.LoadStatus = "Пожалуйста, подождите. Идёт загрузка систем и подсистем...";
            await _vm.InitModelDataAsync();

            _vm.ApplyFilter(null);

            _vm.LoadStatus = "Пожалуйста, подождите. Идёт загрузка данных о ФО систем...";
            await _vm.InitPhysDataAsync();

            _vm.LoadStatus = "Пожалуйста, подождите. Идёт загрузка данных о испытаниях...";
            await _vm.InitTPDataAsync();

            _vm.LoadStatus = "Все данные готовы к визуализации";
            _vm.IsLoading = false;
        }

        // Вызывается, когда контрол удаляется из визуального дерева
        private void OnControlUnloaded(object sender, RoutedEventArgs e)
        {
            // При скрытии/закрытии контрола — сбрасываем все override-цвета
            _vm.RestoreOriginalColors();
        }

        //private void LegendItem_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        //{
        //    if (_vm == null)
        //        return;
        //    _vm.LegendItem_Click(sender, e);
        //}

        private void Systems_Page_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (_vm == null) return;
            _vm.OnClearAll();
            _vm.ReturnVisibilityForAll();
        }

        private void Monitoring_Page_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (_vm == null) return;
            _vm.PerformRestoreVisibility_Monitoring();
            _vm.PerformClearAll_Monitoring();
        }

        private void Passport_Page_MouseLeftButtonDown(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (_vm == null) return;
            _vm.PerformRestoreVisibility_Passsports();
            _vm.PerformClearAll_Passports();
        }

        private void TabControl_SelectionChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            TabControl tabControl = sender as TabControl;
            TabItem tabItem = tabControl.SelectedItem as TabItem;
            string page = (string)tabItem.Header;

            if (page == "Системы")
            {
                _vm.ReturnVisibilityForAll();
                _vm.GrayAllElements();
                _vm.OnClearAll();
                _vm.ReturnVisibilityForAll();
            }
            else if (page == "Мониторинг")
            {
                _vm.ReturnVisibilityForAll();
                _vm.GrayAllElements();
                _vm.PerformRestoreVisibility_Monitoring();
                _vm.PerformClearAll_Monitoring();
            }
            else if (page == "Паспорта")
            {
                _vm.ReturnVisibilityForAll();
                _vm.GrayAllElements();
                _vm.PerformRestoreVisibility_Passsports();
                _vm.PerformClearAll_Passports();
            }
            else if (page == "Данные")
            {

            }
        }
    }
}
