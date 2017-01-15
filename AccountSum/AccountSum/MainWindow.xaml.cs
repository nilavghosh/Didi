using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace AccountSum
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

        private void selectInputFile_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();



            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "XLSX Files (*.xlsx)|*.xlsx|XLS Files (*.xls)|*.xls";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                inputFile.Text = filename;
            }
        }

        private void selectOutputFile_Click(object sender, RoutedEventArgs e)
        {
            // Create OpenFileDialog 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();



            // Set filter for file extension and default file extension 
            dlg.DefaultExt = ".xlsx";
            dlg.Filter = "XLSX Files (*.xlsx)|*.xlsx|XLS Files (*.xls)|*.xls";


            // Display OpenFileDialog by calling ShowDialog method 
            Nullable<bool> result = dlg.ShowDialog();


            // Get the selected file name and display in a TextBox 
            if (result == true)
            {
                // Open document 
                string filename = dlg.FileName;
                outputFile.Text = filename;
            }
        }


        private void convertButton_Click(object sender, RoutedEventArgs e)
        {
            var inputFilePath = inputFile.Text;
            var outputFilePath = outputFile.Text;

            var inpfile = new FileInfo(inputFilePath);
            var outfile = new FileInfo(outputFilePath);

            using (var outModel = new ExcelPackage(outfile))
            {
                ExcelWorkbook outWorkBook = outModel.Workbook;
                if (outWorkBook.Worksheets.Count > 0)
                {
                    var outSheet = outWorkBook.Worksheets["Budget C&F"];
                    var accntCells = outSheet.Cells["C7:C80"].ToList();
                    var catCells = outSheet.Cells["D7:D80"];

                    using (var inpModel = new ExcelPackage(inpfile))
                    {
                        ExcelWorkbook inpWorkBook = inpModel.Workbook;
                        if (inpWorkBook.Worksheets.Count > 0)
                        {

                            var inpSheet = inpWorkBook.Worksheets["Sheet1"];
                            int endRow = inpSheet.Dimension.End.Row;

                            var inpProjNameCells = inpSheet.Cells["A0:A" + endRow.ToString()];
                            var inpEUCells = inpSheet.Cells["B0:B" + endRow.ToString()];
                            var inpAccntCells = inpSheet.Cells["C0:C" + endRow.ToString()];
                            var inpChargeCells = inpSheet.Cells["D0:D" + endRow.ToString()];
                            List<string> loglines = new List<string>();
                            foreach (var accnt in accntCells)
                            {
                                double sum = 0.0;
                                string line = ""; 
                                if (accnt.Value != null && accnt.Value.ToString() != "")
                                {
                                    line += accnt.Value.ToString() + ",";
                                    var accntparts = accnt.Value.ToString().Trim().Split(new char[] { '-' });
                                    string accntNo = accntparts.Count() == 2 ? accntparts[1] : accntparts[0];
                                    var matchingaccnt = inpAccntCells.ToList().Where(cell => cell.Value != null).Where(cell => cell.Value.ToString().Trim().Contains(accntNo) == true);
                                    foreach (var cell in matchingaccnt)
                                    {
                                        //line += cell.Value.ToString() + ",";
                                        if (inpEUCells["B" + cell.Start.Row.ToString()].Value.ToString() == "TX-CF(Dubugue Works)")
                                        {
                                            if (catCells["D" + accnt.Start.Row.ToString()].Value!= null && catCells["D" + accnt.Start.Row.ToString()].Value.ToString().Contains("Sys") == true)
                                            {
                                                line += catCells["D" + accnt.Start.Row.ToString()].Value.ToString() + ",";
                                                if (inpProjNameCells["A" + cell.Start.Row.ToString()].Value.ToString().EndsWith("Systems") == true)
                                                {
                                                    sum += double.Parse(inpChargeCells["D" + cell.Start.Row.ToString()].Value.ToString());
                                                    line += "D" + cell.Start.Row.ToString() + String.Format("({0})", inpChargeCells["D" + cell.Start.Row.ToString()].Value.ToString()) +",";
                                                }
                                            }
                                            else
                                            {
                                                line += "Others" + ",";
                                                if (inpProjNameCells["A" + cell.Start.Row.ToString()].Value.ToString().EndsWith("Systems") != true)
                                                {
                                                    sum += double.Parse(inpChargeCells["D" + cell.Start.Row.ToString()].Value.ToString());
                                                    line += "D" + cell.Start.Row.ToString() + String.Format("({0})", inpChargeCells["D" + cell.Start.Row.ToString()].Value.ToString()) + ",";
                                                }
                                            }
                                        }
                                    }
                                    line += String.Format("Sum={0}", sum.ToString());
                                    loglines.Add(line);
                                    outSheet.SetValue(columnInp.Text.Trim() + accnt.Start.Row.ToString(), sum);
                                }
                            }
                            System.IO.File.WriteAllLines("accntlog.txt", loglines);
                        }
                    }
                    outModel.Save();
                    MessageBox.Show("Done! Please Check.");
                }
            }

        }
    }
}
