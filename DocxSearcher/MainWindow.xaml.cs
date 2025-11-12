using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Media;
using DocumentFormat.OpenXml.Packaging;

namespace DocxSearcher
{
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
            FolderPath.Text = AppContext.BaseDirectory;
        }

        // Безопасная рекурсивная функция для поиска .docx без системных папок
        private static IEnumerable<string> SafeEnumerateDocxFiles(string rootFolder)
        {
            var files = new List<string>();

            try
            {
                foreach (var file in Directory.GetFiles(rootFolder, "*.docx", SearchOption.TopDirectoryOnly))
                    files.Add(file);

                foreach (var dir in Directory.GetDirectories(rootFolder))
                {
                    try
                    {
                        var attr = File.GetAttributes(dir);
                        if ((attr & FileAttributes.System) != 0 ||
                            (attr & FileAttributes.ReparsePoint) != 0)
                            continue;

                        files.AddRange(SafeEnumerateDocxFiles(dir));
                    }
                    catch { } // Игнорируем недоступные папки
                }
            }
            catch { } // Игнорируем корневые недоступные папки

            return files;
        }

        private async void SearchButton_Click(object sender, RoutedEventArgs e)
        {
            string? searchText = SearchText.Text?.Trim();

            if (string.IsNullOrEmpty(searchText))
            {
                System.Windows.MessageBox.Show("Поле поиска не может быть пустым.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            string folderPath = string.IsNullOrWhiteSpace(FolderPath.Text)
                ? AppContext.BaseDirectory
                : FolderPath.Text;

            if (!Directory.Exists(folderPath))
            {
                System.Windows.MessageBox.Show("Указанный путь не существует.", "Ошибка", MessageBoxButton.OK, MessageBoxImage.Error);
                return;
            }

            SearchButton.IsEnabled = false;
            ResultsListBox.Items.Clear();
            StatusLabel.Content = "Подготовка...";
            ProgressBar.Value = 0;
            ProgressBar.Foreground = new SolidColorBrush(Colors.LimeGreen);

            string[] files;
            try
            {
                files = (SearchRecursively.IsChecked == true)
                    ? SafeEnumerateDocxFiles(folderPath).ToArray()
                    : Directory.GetFiles(folderPath, "*.docx", SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                System.Windows.MessageBox.Show($"Ошибка: {ex.Message}");
                SearchButton.IsEnabled = true;
                return;
            }

            if (files.Length == 0)
            {
                StatusLabel.Content = "Файлы .docx не найдены";
                ProgressBar.Value = 100;
                ProgressBar.Foreground = new SolidColorBrush(Colors.Red);
                SearchButton.IsEnabled = true;
                return;
            }

            var foundFiles = new List<string>();

            ProgressBar.Minimum = 0;
            ProgressBar.Maximum = files.Length;

            await Task.Run(() =>
            {
                int current = 0;
                foreach (var file in files)
                {
                    current++;
                    bool contains = SearchInDocx(file, searchText);

                    Dispatcher.Invoke(() =>
                    {
                        ProgressBar.Value = current;
                        StatusLabel.Content = $"Проверка {current} из {files.Length}: {Path.GetFileName(file)}";
                    });

                    if (contains)
                        foundFiles.Add(file);
                }
            });

            ResultsListBox.Items.Clear();
            if (foundFiles.Count > 0)
            {
                foreach (var f in foundFiles)
                    ResultsListBox.Items.Add(f);
                StatusLabel.Content = $"Найдено {foundFiles.Count} файлов";
                ProgressBar.Foreground = new SolidColorBrush(Colors.LimeGreen);
            }
            else
            {
                StatusLabel.Content = "Файлы с искомым текстом не найдены";
                ProgressBar.Foreground = new SolidColorBrush(Colors.Red);
            }

            SearchButton.IsEnabled = true;
        }

        private bool SearchInDocx(string path, string searchText)
        {
            try
            {
                using (var doc = WordprocessingDocument.Open(path, false))
                {
                    var text = doc.MainDocumentPart?.Document?.InnerText ?? "";
                    return text.Contains(searchText, StringComparison.OrdinalIgnoreCase);
                }
            }
            catch
            {
                return false;
            }
        }

        private void BrowseButton_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new System.Windows.Forms.FolderBrowserDialog();
            var result = dialog.ShowDialog();
            if (result == System.Windows.Forms.DialogResult.OK)
                FolderPath.Text = dialog.SelectedPath;
        }

        private void ResultsListBox_MouseDoubleClick(object sender, System.Windows.Input.MouseButtonEventArgs e)
        {
            if (ResultsListBox.SelectedItem == null)
                return;

            string filePath = ResultsListBox.SelectedItem.ToString() ?? "";
            try
            {
                System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                {
                    FileName = filePath,
                    UseShellExecute = true
                });

                StatusLabel.Content = $"Открыт файл: {Path.GetFileName(filePath)}";
            }
            catch (Exception ex)
            {
                StatusLabel.Content = $"Не удалось открыть файл: {ex.Message}";
            }
        }
    }
}
