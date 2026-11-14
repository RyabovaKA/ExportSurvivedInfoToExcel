using System;
using System.Collections.Generic;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using Ascon.Pilot.SDK;
using Microsoft.Win32;

namespace PilotExcelExporter
{
    public class ExportOptionsDialog : Window
    {
        private readonly TextBox _folderTextBox;
        private readonly RadioButton _separateFilesRadio;
        private readonly RadioButton _singleFileRadio;
        private readonly CheckBox _includeServiceCheckBox;
        private readonly CheckBox _resolveRefObjectsCheckBox;

        public string SelectedFolder { get; private set; }
        public bool SeparateFilePerType { get; private set; }
        public bool IncludeServiceAttributes { get; private set; }
        public bool ResolveReferenceObjects { get; private set; }

        public ExportOptionsDialog()
        {
            Title = "Параметры выгрузки в Excel";
            Width = 500;
            Height = 290;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            WindowStyle = WindowStyle.ToolWindow;

            var grid = new Grid { Margin = new Thickness(15) };
            for (int i = 0; i < 7; i++)
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Content = grid;

            var lblFolder = new TextBlock { Text = "Папка для сохранения файлов Excel:", Margin = new Thickness(0, 0, 0, 5) };
            Grid.SetRow(lblFolder, 0);
            grid.Children.Add(lblFolder);

            var folderPanel = new Grid { Margin = new Thickness(0, 0, 0, 12) };
            folderPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = new GridLength(1, GridUnitType.Star) });
            folderPanel.ColumnDefinitions.Add(new ColumnDefinition { Width = GridLength.Auto });

            string defaultFolder = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.DesktopDirectory),
                $"PilotExport_{DateTime.Now:yyyy-MM-dd_HH-mm}");

            _folderTextBox = new TextBox { Text = defaultFolder, Padding = new Thickness(4), Margin = new Thickness(0, 0, 8, 0) };
            Grid.SetColumn(_folderTextBox, 0);
            folderPanel.Children.Add(_folderTextBox);

            var browseBtn = new Button { Content = "Обзор...", Width = 80, Padding = new Thickness(4) };
            browseBtn.Click += BrowseBtn_Click;
            Grid.SetColumn(browseBtn, 1);
            folderPanel.Children.Add(browseBtn);

            Grid.SetRow(folderPanel, 1);
            grid.Children.Add(folderPanel);

            _separateFilesRadio = new RadioButton
            {
                Content = "Отдельный файл Excel (.xlsx) для каждого типа объектов",
                IsChecked = true,
                Margin = new Thickness(0, 0, 0, 6)
            };
            Grid.SetRow(_separateFilesRadio, 2);
            grid.Children.Add(_separateFilesRadio);

            _singleFileRadio = new RadioButton
            {
                Content = "Один файл Excel (отдельный лист на каждый тип объектов)",
                Margin = new Thickness(0, 0, 0, 12)
            };
            Grid.SetRow(_singleFileRadio, 3);
            grid.Children.Add(_singleFileRadio);

            _resolveRefObjectsCheckBox = new CheckBox
            {
                Content = "Подгружать названия объектов из справочников (вместо GUID)",
                IsChecked = true,
                Margin = new Thickness(0, 0, 0, 6)
            };
            Grid.SetRow(_resolveRefObjectsCheckBox, 4);
            grid.Children.Add(_resolveRefObjectsCheckBox);

            _includeServiceCheckBox = new CheckBox
            {
                Content = "Включать служебные атрибуты карточки",
                IsChecked = false,
                Margin = new Thickness(0, 0, 0, 15)
            };
            Grid.SetRow(_includeServiceCheckBox, 5);
            grid.Children.Add(_includeServiceCheckBox);

            var btnPanel = new StackPanel { Orientation = Orientation.Horizontal, HorizontalAlignment = HorizontalAlignment.Right };
            var okBtn = new Button { Content = "Выгрузить", Width = 95, Height = 26, Margin = new Thickness(0, 0, 10, 0), IsDefault = true };
            okBtn.Click += (s, e) =>
            {
                if (string.IsNullOrWhiteSpace(_folderTextBox.Text))
                {
                    MessageBox.Show("Укажите папку для сохранения.", "Внимание", MessageBoxButton.OK, MessageBoxImage.Warning);
                    return;
                }

                SelectedFolder = _folderTextBox.Text.Trim();
                SeparateFilePerType = _separateFilesRadio.IsChecked == true;
                IncludeServiceAttributes = _includeServiceCheckBox.IsChecked == true;
                ResolveReferenceObjects = _resolveRefObjectsCheckBox.IsChecked == true;
                DialogResult = true;
            };

            var cancelBtn = new Button { Content = "Отмена", Width = 85, Height = 26, IsCancel = true };
            btnPanel.Children.Add(okBtn);
            btnPanel.Children.Add(cancelBtn);
            Grid.SetRow(btnPanel, 6);
            grid.Children.Add(btnPanel);
        }

        private void BrowseBtn_Click(object sender, RoutedEventArgs e)
        {
            // Выбор папки стандартным диалогом WPF без необходимости подключать System.Windows.Forms
            var dlg = new SaveFileDialog
            {
                Title = "Выберите папку для сохранения выгрузки",
                FileName = "Выбор папки",
                Filter = "Папка|*.folder",
                CheckFileExists = false,
                CheckPathExists = true,
                ValidateNames = false
            };

            if (Directory.Exists(_folderTextBox.Text))
                dlg.InitialDirectory = _folderTextBox.Text;

            if (dlg.ShowDialog(this) == true)
            {
                string dir = Path.GetDirectoryName(dlg.FileName);
                if (!string.IsNullOrEmpty(dir))
                    _folderTextBox.Text = dir;
            }
        }
    }

    public class ExportProgressWindow : Window
    {
        private readonly IObjectsRepository _objectsRepository;
        private readonly List<Guid> _startIds;
        private readonly string _outputFolder;
        private readonly bool _separateFilePerType;
        private readonly bool _includeServiceAttributes;
        private readonly bool _resolveReferenceObjects;
        private readonly CancellationTokenSource _cts = new CancellationTokenSource();

        private readonly TextBlock _statusText;
        private readonly TextBlock _detailsText;
        private readonly ProgressBar _progressBar;
        private readonly Button _cancelButton;

        public int ExportedFilesCount { get; private set; }
        public int TotalLoadedObjects { get; private set; }
        public bool WasCancelled { get; private set; }
        public string ErrorMessage { get; private set; }

        public ExportProgressWindow(
            IObjectsRepository objectsRepository,
            List<Guid> startIds,
            string outputFolder,
            bool separateFilePerType,
            bool includeServiceAttributes,
            bool resolveReferenceObjects)
        {
            _objectsRepository = objectsRepository;
            _startIds = startIds;
            _outputFolder = outputFolder;
            _separateFilePerType = separateFilePerType;
            _includeServiceAttributes = includeServiceAttributes;
            _resolveReferenceObjects = resolveReferenceObjects;

            Title = "Выгрузка в Excel...";
            Width = 440;
            Height = 175;
            WindowStartupLocation = WindowStartupLocation.CenterScreen;
            ResizeMode = ResizeMode.NoResize;
            WindowStyle = WindowStyle.ToolWindow;

            var grid = new Grid { Margin = new Thickness(15) };
            for (int i = 0; i < 4; i++)
                grid.RowDefinitions.Add(new RowDefinition { Height = GridLength.Auto });
            Content = grid;

            _statusText = new TextBlock { Text = "Инициализация выгрузки...", FontWeight = FontWeights.SemiBold, Margin = new Thickness(0, 0, 0, 6) };
            Grid.SetRow(_statusText, 0);
            grid.Children.Add(_statusText);

            _detailsText = new TextBlock { Text = "Подготовка...", Margin = new Thickness(0, 0, 0, 10), TextTrimming = TextTrimming.CharacterEllipsis };
            Grid.SetRow(_detailsText, 1);
            grid.Children.Add(_detailsText);

            _progressBar = new ProgressBar { IsIndeterminate = true, Height = 18, Margin = new Thickness(0, 0, 0, 14) };
            Grid.SetRow(_progressBar, 2);
            grid.Children.Add(_progressBar);

            _cancelButton = new Button { Content = "Отмена", Width = 90, Height = 26, HorizontalAlignment = HorizontalAlignment.Right };
            _cancelButton.Click += (s, e) =>
            {
                _cancelButton.IsEnabled = false;
                _statusText.Text = "Отмена операции...";
                _cts.Cancel();
            };
            Grid.SetRow(_cancelButton, 3);
            grid.Children.Add(_cancelButton);

            Loaded += async (s, e) => await RunExportAsync();
            Closing += (s, e) => _cts.Cancel();
        }

        private async Task RunExportAsync()
        {
            try
            {
                var uiDispatcher = Dispatcher;
                var progress = new Progress<ExportProgressInfo>(info =>
                {
                    _statusText.Text = info.Stage;
                    _detailsText.Text = info.Details;
                });

                var exporter = new PilotExcelExporter(_objectsRepository, uiDispatcher);
                var result = await exporter.ExportAsync(
                    _startIds,
                    _outputFolder,
                    _separateFilePerType,
                    _includeServiceAttributes,
                    _resolveReferenceObjects,
                    progress,
                    _cts.Token);

                ExportedFilesCount = result.FilesCreated;
                TotalLoadedObjects = result.ObjectsExported;
            }
            catch (OperationCanceledException)
            {
                WasCancelled = true;
            }
            catch (Exception ex)
            {
                ErrorMessage = ex.Message;
            }
            finally
            {
                Close();
            }
        }
    }

    public class ExportProgressInfo
    {
        public string Stage { get; set; }
        public string Details { get; set; }
    }
}