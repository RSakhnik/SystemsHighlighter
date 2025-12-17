using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;

namespace SystemsHighlighter.Tools
{
    /// <summary>
    /// Логика взаимодействия для ProgressDialog.xaml
    /// </summary>
    public partial class ProgressDialog : Window
    {
        private readonly CancellationTokenSource _cts = new CancellationTokenSource();

        public ProgressDialog()
        {
            InitializeComponent();
        }

        public CancellationToken Token => _cts.Token;

        public void SetIndeterminate(string phase)
        {
            Bar.IsIndeterminate = true;
            StatusText.Text = phase;
        }

        public void SetProgress(string phase, int current, int total)
        {
            Bar.IsIndeterminate = false;
            Bar.Value = total > 0 ? (double)current / total * 100.0 : 0;
            StatusText.Text = total > 0
                ? $"{phase} ({current}/{total})"
                : phase;
        }

        private void Cancel_Click(object sender, RoutedEventArgs e)
        {
            CancelBtn.IsEnabled = false;
            _cts.Cancel();
            StatusText.Text = "Отмена...";
        }
    }

}
