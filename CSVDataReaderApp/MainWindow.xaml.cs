using ExcelDataReader;
using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Text;
using System.IO;
using System.Windows;

namespace CSVDataReaderApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private DataSet ds;
        public MainWindow()
        {
            InitializeComponent();
        }

        private void SelectFileBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            if (openFileDialog.ShowDialog() == true)
                FilePathTxt.Text = openFileDialog.FileName;
        }

        private void ProcessFileBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Fjern eksisterende data fra WPF grid view
                CSVDataGrid.ItemsSource = null;
                CSVDataGrid.Items.Refresh();

                // Enc table should probably be set in app config, see note at https://github.com/ExcelDataReader/ExcelDataReader
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                using var stream = new FileStream(FilePathTxt.Text, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);

                var sw = new Stopwatch();
                sw.Start();

                // Indlæs Ecxel binary eller Excel OpenXml filformater
                // using IExcelDataReader reader = ExcelReaderFactory.CreateReader(stream);
                // Indlæs CSV filformat
                using IExcelDataReader reader = ExcelReaderFactory.CreateCsvReader(stream);

                var openTiming = sw.ElapsedMilliseconds;
                
                ds = reader.AsDataSet(new ExcelDataSetConfiguration()
                {
                    UseColumnDataType = false,
                    ConfigureDataTable = (tableReader) => new ExcelDataTableConfiguration()
                    {
                        UseHeaderRow = (bool)FirstRowNamesCheckBox.IsChecked
                    }
                });

                StatusStripTxt.Text = "Elapsed: " + sw.ElapsedMilliseconds.ToString() + " ms (" + openTiming.ToString() + " ms to open)";

                // her kan der ex findes 10 Columns på ds.Tables[0].Columns ved load af data fra brancheforeningen
                var tablenames = GetTablenames(ds.Tables);
                ExcelSheets.ItemsSource = tablenames;

                // Case: Reload til grid af ny csv datafil
                if (ExcelSheets.SelectedIndex == 0)
                    SelectTable();
                // Case: Første gang der vælges en csv datafil
                else if (tablenames.Count > 0)
                    ExcelSheets.SelectedIndex = 0;
                
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString(), ex.Message, MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private static IList<string> GetTablenames(DataTableCollection tables)
        {
            var tableList = new List<string>();
            foreach (var table in tables)
            {
                tableList.Add(table.ToString());
            }

            return tableList;
        }

        private void SelectTable()
        {
            var tablename = ExcelSheets.SelectedItem.ToString();           
            CSVDataGrid.AutoGenerateColumns = true;
            CSVDataGrid.ItemsSource = ds.Tables[0].AsDataView(); 
            // new DataView(ds.Tables[0]); //ds.Tables[0].AsDataView(); 
            // AsEnumerable(); //.Rows; //.AsDataView(); // dataset            
            CSVDataGrid.DataContext = ds.Tables[0];
            // GetValues(ds, tablename);
        }

        public static void GetValues(DataSet dataset, string sheetName)
        {
            foreach (DataRow row in dataset.Tables[sheetName].Rows)
            {
                foreach (var value in row.ItemArray)
                {
                    // Console.WriteLine("{0}, {1}", value, value.GetType());
                    var x = value;
                    var y = value.GetType();
                }
            }
        }


        private void OnSelectedIndexChanged(object sender, System.Windows.Controls.SelectionChangedEventArgs e)
        {
            SelectTable();
        }
    }
}