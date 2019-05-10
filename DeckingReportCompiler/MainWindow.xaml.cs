using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;

using Excel = Microsoft.Office.Interop.Excel;

using Microsoft.Win32;

namespace DeckingReportCompiler
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();

            clearanceFilePath_TextChanged(null, null);
        }

        private void browseClearanceFile_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "Clearance Results (*.txt)|*.txt";
            openFileDialog.Title = "Open Clearance Results";

            if (openFileDialog.ShowDialog() == true)
                clearanceFilePath.Text = openFileDialog.FileName;
        }

        private void compileButton_Click(object sender, RoutedEventArgs e)
        {
            var pathToClearanceFile = clearanceFilePath.Text;
            string pathToResultsFile = null;
            string pathToWorstCaseClearanceFile = null;

            var directoryName = Path.GetDirectoryName(pathToClearanceFile);
            var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(pathToClearanceFile);

            var saveExcelFileDialog = new SaveFileDialog();
            //saveExcelFileDialog.Filter = "Excel Macro-Enabled Workbook (*.xlsm)|*.xlsm";
            saveExcelFileDialog.Filter = "Excel Workbook (*.xlsx)|*.xlsx";
            saveExcelFileDialog.OverwritePrompt = true;
            saveExcelFileDialog.InitialDirectory = directoryName;
            saveExcelFileDialog.FileName = fileNameWithoutExtension;
            saveExcelFileDialog.Title = "Save Excel report";

            if (saveExcelFileDialog.ShowDialog() == true)
                pathToResultsFile = saveExcelFileDialog.FileName;
            else
                return;

            var saveWorstCaseClearanceFileDialog = new SaveFileDialog();
            saveWorstCaseClearanceFileDialog.Filter = "Clearance Results (*.txt)|*.txt";
            saveWorstCaseClearanceFileDialog.OverwritePrompt = true;
            saveWorstCaseClearanceFileDialog.InitialDirectory = directoryName;
            saveWorstCaseClearanceFileDialog.FileName = "WORST CASE " + fileNameWithoutExtension;
            saveWorstCaseClearanceFileDialog.Title = "Save worst case Clearance Results";

            if (saveWorstCaseClearanceFileDialog.ShowDialog() == true)
                pathToWorstCaseClearanceFile = saveWorstCaseClearanceFileDialog.FileName;
            else
                return;

            progressOverlay.Visibility = System.Windows.Visibility.Visible;

            var deReC = new DeReC();
            deReC.OnError += message => Console.Error.WriteLine(message);
            deReC.OnCompileProgress += value => progressBar.Dispatcher.Invoke(() => progressBar.Value = value / 2);

            var excelManager = new ExcelManager();
            excelManager.OnProgress += value => progressBar.Dispatcher.Invoke(() => progressBar.Value = 0.5 + value / 2);

            var stepSize = this.stepSize.Value.Value;

            Task.Run(() => deReC.Compile(pathToClearanceFile, stepSize))
                .ContinueWith(task =>
                {
                    var clearanceFileMetaData = task.Result;

                    Task.WaitAll(new Task[]
                    {
                        Task.Factory.StartNew(() =>
                        {
                            var partPairs = clearanceFileMetaData.PartPairs;

                            if (partPairs == null) return;

                            using (var excelFileStream = new FileStream(pathToResultsFile, FileMode.Create, FileAccess.Write))
                            {
                                excelManager.Write(partPairs, excelFileStream);
                            }
                        }),

                        Task.Factory.StartNew(() =>
                        {
                            deReC.SaveWorstCaseClearanceFile(clearanceFileMetaData, pathToWorstCaseClearanceFile);
                        })
                    });
                })
                .ContinueWith(task => completeOverlay.Dispatcher.Invoke(() => completeOverlay.Visibility = System.Windows.Visibility.Visible));
        }

        private void doneButton_Click(object sender, RoutedEventArgs e)
        {
            progressOverlay.Visibility = System.Windows.Visibility.Hidden;
            completeOverlay.Visibility = System.Windows.Visibility.Hidden;
        }

        private void clearanceFilePath_TextChanged(object sender, TextChangedEventArgs e)
        {
            var pathToClearanceFile = clearanceFilePath.Text;

            if (File.Exists(pathToClearanceFile) && Path.GetExtension(pathToClearanceFile).ToLower() == ".txt")
            {
                compileButton.IsEnabled = true;
                clearanceFilePath.Background = null;
            }

            else
            {
                compileButton.IsEnabled = false;
                clearanceFilePath.Background = new SolidColorBrush((System.Windows.Media.Color)System.Windows.Media.ColorConverter.ConvertFromString("#E5F9B9B9"));
            }
        }
    }
}