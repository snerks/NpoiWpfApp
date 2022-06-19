using NPOI.HSSF.UserModel;
using NPOI.POIFS.FileSystem;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NpoiWpfApp.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace NpoiWpfApp
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void okButton_Click(object sender, RoutedEventArgs e)
        {
            var extractFile = @"C:\D\Samples\DotNet31\NpoiWpfApp\NpoiWpfApp\Book1.xlsx";

            // , FileMode.Create, FileAccess.Write

            //using (FileStream file = File.Open(extractFile, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            //{
            //    IWorkbook workbook = new XSSFWorkbook(file);

            //    if (null == workbook)
            //    {
            //        Console.WriteLine(string.Format("Excel Workbook '{0}' could not be opened.", extractFile));
            //        return;
            //    }

            //    var sheetName = "Sheet1";

            //    ISheet sheet = workbook.GetSheet(sheetName);

            //}

            var mapper = new Npoi.Mapper.Mapper(extractFile);
            mapper.FirstRowIndex = 1;

            ISheet currentSheet = mapper.Workbook.GetSheetAt(mapper.Workbook.ActiveSheetIndex);

            List<string> columnsHeader = 
                currentSheet
                    .GetRow(currentSheet.FirstRowNum)
                    .Cells
                    .Select(columnHeader => columnHeader.StringCellValue)
                    .ToList();

            var items = mapper.Take<SampleClass>("Sheet3").ToList();

            //var errorCount = 0;

            //foreach (var item in items)
            //{
            //    if (item.ErrorColumnIndex > -1)
            //    {
            //        errorCount++;
            //    }
            //}

            var haveErrors = items.Any(i => i.ErrorColumnIndex > -1);

            var itemValues = new List<SampleClass>();

            if (!haveErrors)
            {
                itemValues = items.Select(i => i.Value).ToList();
            }

            resultsDataGrid.ItemsSource = itemValues;
        }
    }
}
