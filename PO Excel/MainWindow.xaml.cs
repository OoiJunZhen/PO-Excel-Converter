using Microsoft.Win32;
using System.ComponentModel;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media;
using PO_Excel.Model;

namespace PO_Excel
{
    public partial class MainWindow : Window
    {
        private BackgroundWorker loadWorker;
        private BackgroundWorker convertWorker;
        private List<DataList>? originalDataList;

        private CancellationTokenSource _convertCancellationTokenSource;
        private CancellationTokenSource _loadCancellationTokenSource;

        private bool isSearchHandled = false;

        public MainWindow()
        {
            InitializeComponent();
            InitializeBackgroundWorkers();
        }

        private void Window_StateChanged(object sender, System.EventArgs e)
        {
            if (WindowState == WindowState.Normal)
            {
                CenterWindowOnScreen();
            }
        }

        private void CenterWindowOnScreen()
        {
            var screenWidth = SystemParameters.PrimaryScreenWidth;
            var screenHeight = SystemParameters.PrimaryScreenHeight;

            var windowWidth = this.Width;
            var windowHeight = this.Height;
            this.Left = (screenWidth / 2) - (windowWidth / 2);
            this.Top = (screenHeight / 2) - (windowHeight / 2);
        }

        private void InitializeBackgroundWorkers()
        {
            loadWorker = new BackgroundWorker
            {
                WorkerReportsProgress = true
            };
            loadWorker.DoWork += LoadWorker_DoWork;
            loadWorker.ProgressChanged += LoadWorker_ProgressChanged;
            loadWorker.RunWorkerCompleted += LoadWorker_RunWorkerCompleted;

            convertWorker = new BackgroundWorker
            {
                WorkerReportsProgress = true
            };
            convertWorker.DoWork += ConvertWorker_DoWork;
            convertWorker.ProgressChanged += ConvertWorker_ProgressChanged;
            convertWorker.RunWorkerCompleted += ConvertWorker_RunWorkerCompleted;
        }

        private void TxtFilePath_PreviewDragOver(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                e.Effects = DragDropEffects.Copy;
            }
            else
            {
                e.Effects = DragDropEffects.None;
            }
            e.Handled = true;
        }

        private void TxtFilePath_Drop(object sender, DragEventArgs e)
        {
            if (e.Data.GetDataPresent(DataFormats.FileDrop))
            {
                string[] files = (string[])e.Data.GetData(DataFormats.FileDrop);
                if (files.Length > 0)
                {
                    string droppedFile = files[0];
                    if (IsExcelFile(droppedFile))
                    {
                        if (string.IsNullOrEmpty(txtBox1.Text))
                        {
                            MessageBox.Show("Please enter project code.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                            return;
                        }
                        if (sender == txtFilePath1)
                        {
                            txtFilePath1.Text = droppedFile;
                            txtSearch.Text = string.Empty;
                            originalDataList = null;
                            dataGrid.ItemsSource = null;
                            UpdateStatus("Loading... Please wait.", Brushes.SkyBlue);
                            UpdateStatus1("");
                            var arguments = new Tuple<string, string>(droppedFile, txtBox1.Text);
                            loadWorker.RunWorkerAsync(arguments);

                            btnConvert.IsEnabled = false;
                            btnClear.IsEnabled = false;
                            btnSelect1.IsEnabled = false;
                            btnSelect2.IsEnabled = false;

                            txtFilePath1.AllowDrop = false;
                            txtFilePath2.AllowDrop = false;
                            txtSearch.IsEnabled = false;
                            txtBox1.IsEnabled = false;

                            btnStop.Visibility = Visibility.Visible;
                            btnStop.IsEnabled = true;

                            _loadCancellationTokenSource = new CancellationTokenSource();
                        }
                        else if (sender == txtFilePath2)
                        {
                            txtFilePath2.Text = droppedFile;
                            UpdateStatus("Project List Excel File selected.", Brushes.ForestGreen);
                        }
                    }
                    else
                    {
                        MessageBox.Show("Please drop a valid Excel file.", "Invalid File", MessageBoxButton.OK, MessageBoxImage.Warning);
                    }
                }
            }
        }

        private static bool IsExcelFile(string filePath)
        {
            string? extension = System.IO.Path.GetExtension(filePath)?.ToLower();
            return extension == ".xls" || extension == ".xlsx" || extension == ".xlsm";
        }

        private void TxtFilePath_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (sender is TextBox textBox && txtFilePath1.Text != "Drag Or Select")
            {
                var placeholder = textBox.Template.FindName("Placeholder", textBox) as TextBlock;

                if (textBox.Text.Length > 0)
                {
                    textBox.Foreground = new SolidColorBrush(Colors.White);
                    if (placeholder != null)
                    {
                        placeholder.Visibility = Visibility.Collapsed;
                    }
                }
                else
                {
                    textBox.Foreground = new SolidColorBrush(Colors.Gray);
                    if (placeholder != null)
                    {
                        placeholder.Visibility = Visibility.Visible;
                    }
                }
            }
        }

        private void BtnFile1_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtBox1.Text))
            {
                MessageBox.Show("Please enter project code.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            OpenFileDialog openFileDialog = new()
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm"
            };
            if (openFileDialog.ShowDialog() == true)
            {
                txtFilePath1.Text = openFileDialog.FileName;
                txtSearch.Text = string.Empty;
                originalDataList = null;
                dataGrid.ItemsSource = null;
                UpdateStatus("Loading... Please wait.", Brushes.SkyBlue);
                UpdateStatus1("");
                var arguments = new Tuple<string, string>(openFileDialog.FileName, txtBox1.Text);
                loadWorker.RunWorkerAsync(arguments);

                btnConvert.IsEnabled = false;
                btnClear.IsEnabled = false;
                btnSelect1.IsEnabled = false;
                btnSelect2.IsEnabled = false;

                txtFilePath1.AllowDrop = false;
                txtFilePath2.AllowDrop = false;
                txtSearch.IsEnabled = false;
                txtBox1.IsEnabled = false;


                btnStop.Visibility = Visibility.Visible;
                btnStop.IsEnabled = true;

                
                _loadCancellationTokenSource = new CancellationTokenSource();
            }
        }

        private void BtnFile2_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtBox1.Text))
            {
                MessageBox.Show("Please enter project code.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            OpenFileDialog openFileDialog = new()
            {
                Filter = "Excel Files|*.xls;*.xlsx;*.xlsm",
                Title = "Select Save Location"
            };

            if (openFileDialog.ShowDialog() == true)
            {
                txtFilePath2.Text = openFileDialog.FileName;
                UpdateStatus("Project List Excel File selected.", Brushes.ForestGreen);
            }
        }

        private void BtnConvert_Click(object sender, RoutedEventArgs e)
        {
            if (string.IsNullOrEmpty(txtBox1.Text))
            {
                MessageBox.Show("Please enter project code.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }
            if ((string.IsNullOrEmpty(txtFilePath1.Text) || txtFilePath1.Text.Equals("Drag Or Select")) &&
                (string.IsNullOrEmpty(txtFilePath2.Text) || txtFilePath2.Text.Equals("Drag Or Select")))
            {
                MessageBox.Show("Please select both PR Excel File and Project List Excel File.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(txtFilePath1.Text) || txtFilePath1.Text.Equals("Drag Or Select"))
            {
                MessageBox.Show("Please select PR Excel File.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (string.IsNullOrEmpty(txtFilePath2.Text) || txtFilePath2.Text.Equals("Drag Or Select"))
            {
                MessageBox.Show("Please select Project List Excel File.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            string projectFileName = System.IO.Path.GetFileNameWithoutExtension(txtFilePath2.Text);
            string[] parts = projectFileName.Split(" - ", StringSplitOptions.RemoveEmptyEntries);

            if (parts.Length > 0 && !parts[0].Contains(txtBox1.Text))
            {
                MessageBox.Show($"The project list file ({parts[0]}) does not match the entered project code ({txtBox1.Text}).", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            if (dataGrid.ItemsSource is not List<DataList> dataList || dataList.Count == 0)
            {
                UpdateStatus("No data to save.", Brushes.LightCoral);
                MessageBox.Show("There is no data to save.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            UpdateStatus("Converting... Please wait.", Brushes.SkyBlue);
            btnConvert.IsEnabled = false;
            btnClear.IsEnabled = false;
            btnSelect1.IsEnabled = false;
            btnSelect2.IsEnabled = false;

            txtFilePath1.AllowDrop = false;
            txtFilePath2.AllowDrop = false;
            txtSearch.IsEnabled = false;
            txtBox1.IsEnabled = false;

            btnStop.Visibility = Visibility.Visible;
            btnStop.IsEnabled = true;

            _convertCancellationTokenSource = new CancellationTokenSource();

            var arguments = new Tuple<string, string, List<DataList>, string>(txtFilePath1.Text, txtFilePath2.Text, dataList, txtBox1.Text);
            convertWorker.RunWorkerAsync(arguments);
        }


        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            txtSearch.Text = string.Empty;
            txtBox1.Text = string.Empty;
            originalDataList = null;
            dataGrid.ItemsSource = null;
            txtFilePath1.Text = "Drag Or Select";
            txtFilePath2.Text = "Drag Or Select";
            txtFilePath1.Foreground = new SolidColorBrush(Colors.Gray);
            txtFilePath2.Foreground = new SolidColorBrush(Colors.Gray);
            UpdateStatus("STATUS", new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2A2B3D")));
            UpdateStatus1("");
        }

        private void BtnStop_Click(object sender, RoutedEventArgs e)
        {
            _convertCancellationTokenSource?.Cancel();
            _loadCancellationTokenSource?.Cancel();
        }

        private void LoadWorker_DoWork(object? sender, DoWorkEventArgs e)
        {
            var arguments = e.Argument as Tuple<string, string> ?? throw new ArgumentNullException(null, nameof(e.Argument)); ;
            string filePath = arguments.Item1;
            string textBox1 = arguments.Item2;
            try
            {
                using var excelHelper = new ExcelHelper(filePath);
                {
                    BackgroundWorker? worker = sender as BackgroundWorker;
                    var dataList = excelHelper.LoadColumns(textBox1, worker, _loadCancellationTokenSource.Token);
                    e.Result = dataList;
                }
            }
            catch (Exception ex)
            {
                e.Result = ex.Message;
            }
        }

        private void LoadWorker_ProgressChanged(object? sender, ProgressChangedEventArgs e)
        {
            string progressMessage = e.UserState as string ?? "Working...";

            Brush progressBrush = e.ProgressPercentage switch
            {
                <= 25 => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#87CEEB")),
                <= 50 => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4A90E2")),
                <= 75 => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1565C0")),
                <= 100 => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1A66B2")),
                _ => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2A2B3D")),
            };


            UpdateStatus($"{progressMessage} ({e.ProgressPercentage}%)", progressBrush);
        }

        private void LoadWorker_RunWorkerCompleted(object? sender, RunWorkerCompletedEventArgs e)
        {
            btnConvert.IsEnabled = true;
            btnClear.IsEnabled = true;
            btnSelect1.IsEnabled = true;
            btnSelect2.IsEnabled = true;
            txtFilePath1.AllowDrop = true;
            txtFilePath2.AllowDrop = true;
            txtSearch.IsEnabled = true;
            txtBox1.IsEnabled = true;
            btnStop.Visibility = Visibility.Collapsed;

            if (e.Cancelled)
            {
                UpdateStatus("Loading was canceled.", Brushes.Orange);
                MessageBox.Show("The loading was canceled.", "Canceled", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else if (e.Result is string errorMessage)
            {
                UpdateStatus("Error loading file.", Brushes.LightCoral);
                MessageBox.Show($"An error occurred while loading columns:\n{errorMessage}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            else
            {
                originalDataList = e.Result as List<DataList>;
                dataGrid.ItemsSource = originalDataList;
                UpdateStatus("Data loaded complete.", Brushes.ForestGreen);
                UpdateStatus1($"{originalDataList?.Count}");
            }
        }

        private void ConvertWorker_DoWork(object? sender, DoWorkEventArgs e)
        {
            var arguments = e.Argument as Tuple<string, string, List<DataList>, string> ?? throw new ArgumentNullException(nameof(e.Argument), "Invalid arguments.");
            string loadFilePath = arguments.Item1;
            string saveFilePath = arguments.Item2;
            var dataList = arguments.Item3;
            string textBox1 = arguments.Item4;

            BackgroundWorker? worker = sender as BackgroundWorker;

            try
            {
                using var excelHelper = new ExcelHelper(loadFilePath);
                excelHelper.SaveToFile(dataList, saveFilePath, worker, textBox1, _convertCancellationTokenSource.Token);
                e.Result = "Conversion complete.";
            }
            catch (Exception ex)
            {
                e.Result = ex.Message;
            }
        }



        private void ConvertWorker_ProgressChanged(object? sender, ProgressChangedEventArgs e)
        {
            string progressMessage = e.UserState as string ?? "Working...";

            Brush progressBrush = e.ProgressPercentage switch
            {
                <= 25 => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#87CEEB")),
                <= 50 => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#4A90E2")),
                <= 75 => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1565C0")),
                <= 100 => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#1A66B2")),
                _ => new SolidColorBrush((Color)ColorConverter.ConvertFromString("#2A2B3D")),
            };


            UpdateStatus($"{progressMessage} ({e.ProgressPercentage}%)", progressBrush);
        }

        private void ConvertWorker_RunWorkerCompleted(object? sender, RunWorkerCompletedEventArgs e)
        {
            btnConvert.IsEnabled = true;
            btnClear.IsEnabled = true;
            btnSelect1.IsEnabled = true;
            btnSelect2.IsEnabled = true;
            txtFilePath1.AllowDrop = true;
            txtFilePath2.AllowDrop = true;
            txtSearch.IsEnabled = true;
            txtBox1.IsEnabled = true;

            btnStop.Visibility = Visibility.Collapsed;

            if (e.Cancelled)
            {
                UpdateStatus("Conversion was canceled.", Brushes.Orange);
                MessageBox.Show("The conversion was canceled.", "Canceled", MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else if (e.Result is string resultMessage)
            {
                if (resultMessage == "Conversion complete.")
                {
                    UpdateStatus(resultMessage, Brushes.ForestGreen);
                    MessageBox.Show($"Conversion complete!", "Success", MessageBoxButton.OK, MessageBoxImage.Information);
                }
                else
                {
                    UpdateStatus("Error converting file.", Brushes.LightCoral);
                    MessageBox.Show($"An error occurred during conversion:\n{resultMessage}", "Error", MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }

        private void TxtSearch_TextChanged(object sender, TextChangedEventArgs e)
        {
            if (isSearchHandled) return;

            if (originalDataList == null || originalDataList.Count == 0)
            {
                isSearchHandled = true;
                MessageBox.Show("Please load the data before searching.", "Error", MessageBoxButton.OK, MessageBoxImage.Warning);
                txtSearch.TextChanged -= TxtSearch_TextChanged;
                txtSearch.Text = string.Empty;
                txtSearch.TextChanged += TxtSearch_TextChanged;
                isSearchHandled = false;
                return;
            }

            string searchTerm = txtSearch.Text.Trim().ToLower();

            if (string.IsNullOrEmpty(searchTerm))
            {
                dataGrid.ItemsSource = originalDataList;
                UpdateStatus1($"{originalDataList.Count}");
                return;
            }

            var filteredData = originalDataList
                .Where(data => data.PRProjectCode?.ToLower().Contains(searchTerm, StringComparison.CurrentCultureIgnoreCase) == true
                || data.PRNo?.ToLower().Contains(searchTerm, StringComparison.CurrentCultureIgnoreCase) == true
                || data.PRMaterialCode?.ToLower().Contains(searchTerm, StringComparison.CurrentCultureIgnoreCase) == true
                || data.PRApprovedOn?.ToLower().Contains(searchTerm, StringComparison.CurrentCultureIgnoreCase) == true
                || data.POProjectCode?.ToLower().Contains(searchTerm, StringComparison.CurrentCultureIgnoreCase) == true
                || data.PONo?.ToLower().Contains(searchTerm, StringComparison.CurrentCultureIgnoreCase) == true
                || data.POMaterialCode?.ToLower().Contains(searchTerm, StringComparison.CurrentCultureIgnoreCase) == true
                || data.ReceivedQty?.ToLower().Contains(searchTerm, StringComparison.CurrentCultureIgnoreCase) == true
                || data.POApprovedOn?.ToLower().Contains(searchTerm, StringComparison.CurrentCultureIgnoreCase) == true)
                .ToList();

            dataGrid.ItemsSource = filteredData;
            UpdateStatus1($"{filteredData.Count}");
        }

        private void UpdateStatus(string message, Brush background)
        {
            lblStatus.Content = message;
            lblStatus.Background = background;
        }

        private void UpdateStatus1(string message)
        {
            lblStatus1.Content = message;
        }
    }
}