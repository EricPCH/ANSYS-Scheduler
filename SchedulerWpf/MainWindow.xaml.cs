using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using Microsoft.Win32;

namespace SchedulerWpf
{
    public partial class MainWindow : Window
    {
        private readonly ObservableCollection<QueueItem> _queue = new();
        private readonly ObservableCollection<FinishedItem> _finished = new();
        private CancellationTokenSource _cts = new();
        private Process? _currentProcess;
        private bool _isSimulating;
        private readonly string _ansysPath = Environment.GetEnvironmentVariable("ANSYSEDT_PATH") ??
            @"C:\Program Files\ANSYS Inc\v251\AnsysEM\ansysedt";

        public ObservableCollection<QueueItem> Queue => _queue;
        public ObservableCollection<FinishedItem> Finished => _finished;
        public string StatusText { get; set; } = "Ready";

        public MainWindow()
        {
            InitializeComponent();
            DataContext = this;
        }

        private void AddFile_Click(object sender, RoutedEventArgs e)
        {
            var dlg = new OpenFileDialog
            {
                Filter = "AEDT Files (*.aedt;*.aedtz)|*.aedt;*.aedtz",
                Multiselect = true
            };
            if (dlg.ShowDialog() == true)
            {
                foreach (string file in dlg.FileNames)
                {
                    if (!file.EndsWith(".aedt", StringComparison.OrdinalIgnoreCase) &&
                        !file.EndsWith(".aedtz", StringComparison.OrdinalIgnoreCase))
                    {
                        MessageBox.Show("Only .aedt or .aedtz files are allowed.");
                        continue;
                    }
                    var lockPath = file + ".lock";
                    if (File.Exists(lockPath))
                    {
                        try { File.Delete(lockPath); } catch { }
                    }
                    _queue.Add(new QueueItem
                    {
                        FileName = Path.GetFileName(file),
                        NonGraphical = NgCheckBox.IsChecked == true,
                        SubmitTime = DateTime.Now,
                        FullPath = file
                    });
                }
                if (!_isSimulating)
                    _ = StartSimulationAsync();
            }
        }

        private void Remove_Click(object sender, RoutedEventArgs e)
        {
            var selected = QueueGrid.SelectedItems.Cast<QueueItem>().ToList();
            if (selected.Count == 0) return;
            foreach (var item in selected)
            {
                if (_isSimulating && ReferenceEquals(item, _queue.First()))
                {
                    MessageBox.Show("Cannot remove file being simulated.");
                    continue;
                }
                _queue.Remove(item);
            }
        }

        private void Up_Click(object sender, RoutedEventArgs e)
        {
            if (QueueGrid.SelectedItem is QueueItem item)
            {
                int index = _queue.IndexOf(item);
                if (index > 0 && !(_isSimulating && index - 1 == 0))
                {
                    _queue.Move(index, index - 1);
                }
            }
        }

        private void Down_Click(object sender, RoutedEventArgs e)
        {
            if (QueueGrid.SelectedItem is QueueItem item)
            {
                int index = _queue.IndexOf(item);
                if (index < _queue.Count - 1 && !( _isSimulating && index == 0))
                {
                    _queue.Move(index, index + 1);
                }
            }
        }

        private async Task StartSimulationAsync()
        {
            _isSimulating = true;
            StatusText = "Running";
            while (_queue.Count > 0 && !_cts.IsCancellationRequested)
            {
                var item = _queue.First();
                StatusText = $"Simulating: {item.FullPath}";
                QueueGrid.Items.Refresh();
                var start = DateTime.Now;
                if (item.FullPath.EndsWith(".aedt", StringComparison.OrdinalIgnoreCase))
                {
                    var args = $"-batchsolve {(item.NonGraphical ? "-ng " : string.Empty)}\"{item.FullPath}\"";
                    var psi = new ProcessStartInfo(_ansysPath, args)
                    {
                        UseShellExecute = false
                    };
                    _currentProcess = Process.Start(psi);
                    if (_currentProcess != null)
                        await _currentProcess.WaitForExitAsync(_cts.Token);
                    _currentProcess = null;
                    if (_cts.IsCancellationRequested) break;
                }
                else
                {
                    try { await Task.Delay(1000, _cts.Token); } catch { }
                }
                var stop = DateTime.Now;
                _queue.RemoveAt(0);
                _finished.Add(new FinishedItem
                {
                    FileName = item.FileName,
                    StartTime = start,
                    StopTime = stop,
                    Duration = stop - start,
                    FullPath = item.FullPath
                });
            }
            StatusText = "Finished";
            _isSimulating = false;
            QueueGrid.Items.Refresh();
        }

        protected override void OnClosing(System.ComponentModel.CancelEventArgs e)
        {
            _cts.Cancel();
            try
            {
                _currentProcess?.Kill(true);
            }
            catch { }
            base.OnClosing(e);
        }
    }

    public class QueueItem
    {
        public string FileName { get; set; } = string.Empty;
        public bool NonGraphical { get; set; }
        public DateTime SubmitTime { get; set; }
        public string FullPath { get; set; } = string.Empty;
    }

    public class FinishedItem
    {
        public string FileName { get; set; } = string.Empty;
        public DateTime StartTime { get; set; }
        public DateTime StopTime { get; set; }
        public TimeSpan Duration { get; set; }
        public string FullPath { get; set; } = string.Empty;
    }
}
