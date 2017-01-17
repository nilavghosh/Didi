using OfficeOpenXml;
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

namespace SpeedNetworking
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
            var inputFilePath = inputFile.Text;
            var outputFilePath = outputFile.Text;

            var inpfile = new FileInfo(inputFilePath);

            using (var inModel = new ExcelPackage(inpfile))
            {
                ExcelWorkbook inWorkBook = inModel.Workbook;

                //for (int i = 1; i < 6; i++)
                //{

                //    ExcelWorksheet Round1 = inWorkBook.Worksheets.Add("Round"+i.ToString());

                //    Round1.Cells["B1"].Value = "Round" + i.ToString();

                //    inModel.Save();
                //}
                ExcelWorksheet InputSheet = inWorkBook.Worksheets["Input"];

                int iColCnt = InputSheet.Dimension.End.Column;

                //Rows minus Header
                int iRowCnt = InputSheet.Dimension.End.Row;

                for (int round = 1; round <= 5; round++)
                {
                    ExcelWorksheet Round = inWorkBook.Worksheets.Add("Round" + round.ToString());
                    Round.Cells["E3"].Value = "LEADER";
                    Round.Cells["F3"].Value = "ROWS";
                    Round.Cells["G3"].Value = "C1";
                    Round.Cells["H3"].Value = "C2";
                    Round.Cells["I3"].Value = "C3";
                    Round.Cells["E4:E" + (iRowCnt + 4).ToString()].Value = InputSheet.Cells["A2:A" + iRowCnt.ToString()].Value;
                }
                try
                {

                    List<IEnumerable<string>> currentRoundOtherOptions = new List<IEnumerable<string>>();
                    for (int leaderRowNo = 1; leaderRowNo < iRowCnt; leaderRowNo++)
                    {
                        ExcelRange leaderRow = InputSheet.Cells["A" + (leaderRowNo + 1).ToString() + ":F" + (leaderRowNo + 1).ToString()];
                        List<String> groupMembers = leaderRow.ToList().Select(cell => cell.Value.ToString()).ToList();

                        var groupCombos = GetKCombs<string>(leaderRow.ToList().Select(cell => cell.Value.ToString()).ToList(), 2).ToList();


                        var leaderCombos = groupCombos.Where(combo => combo.Contains(groupMembers[0])).ToList();
                        for (int roundNo = 1; roundNo <= 5; roundNo++)
                        {
                            var RoundSheet = inWorkBook.Worksheets["Round" + roundNo.ToString()];
                            var currentRoundLeaderCombo = leaderCombos[0].ToList();
                            string cellValue = currentRoundLeaderCombo[0] + "-" + currentRoundLeaderCombo[1];
                            RoundSheet.Cells["G" + (3 + leaderRowNo).ToString()].Value = cellValue;
                            leaderCombos.Remove(leaderCombos[0]);

                            currentRoundOtherOptions = groupCombos.Where(combo => (!cellValue.Contains(combo.ToList()[0]) && !cellValue.Contains(combo.ToList()[1]))).ToList();

                            var roundCombis = GetKCombs<int>(Enumerable.Range(0, currentRoundOtherOptions.Count).ToList(), 2).ToList();
                            var roundSetIndices = Get2IndependentCombinations(currentRoundOtherOptions, roundCombis);

                            cellValue = currentRoundOtherOptions[roundSetIndices[0]].ToList()[0] + "-" + currentRoundOtherOptions[roundSetIndices[0]].ToList()[1];
                            RoundSheet.Cells["H" + (3 + leaderRowNo).ToString()].Value = cellValue;

                            cellValue = currentRoundOtherOptions[roundSetIndices[1]].ToList()[0] + "-" + currentRoundOtherOptions[roundSetIndices[1]].ToList()[1];
                            RoundSheet.Cells["I" + (3 + leaderRowNo).ToString()].Value = cellValue;

                            groupCombos.Remove(currentRoundOtherOptions[roundSetIndices[0]]);
                            groupCombos.Remove(currentRoundOtherOptions[roundSetIndices[1]]);

                        }
                    }
                }
                catch (Exception exp)
                {
                    Console.WriteLine("Error");
                }
                finally
                {
                    inModel.Save();
                }
                try
                {
                    ExcelWorksheet Layout = inWorkBook.Worksheets.Add("Output");
                    var participants = InputSheet.Cells["A2:F11"];
                    int m = 3;
                    int n = 1;
                    for (int row = 2; row <= iRowCnt; row++)
                    {
                        for (int col = 1; col <= iColCnt; col++)
                        {
                            string p = InputSheet.Cells[row, col].Value.ToString();

                            Layout.Cells[m, n].Value = p;
                            for (int r = 1; r <= 5; r++)
                            {
                                n += 2;
                                var RoundSheet = inWorkBook.Worksheets["Round" + r.ToString()];
                                var roundSlots = RoundSheet.Cells["G" + (3 + row-1).ToString() + ":I" + (3 + row-1).ToString()];
                                roundSlots.ToList().ForEach(cell =>
                                {
                                    if (cell.Value.ToString().Contains(p))
                                    {
                                        string partner = cell.Value.ToString().Split('-').Where(part => part != p).ToList().First();
                                        Layout.Cells[m, n].Value = partner;
                                    }
                                });
                            }

                            m += 5;
                            n = 1;
                        }
                    }
               
                }
                catch (Exception exp)
                {
                    Console.WriteLine("ErrorEventArgs");
                }
                finally
                {
                    inModel.Save();
                }
            }

        }

        static IEnumerable<IEnumerable<T>> GetCombinations<T>(IEnumerable<T> list, int length)
        {
            if (length == 1) return list.Select(t => new T[] { t });

            return GetCombinations(list, length - 1)
                .SelectMany(t => list, (t1, t2) => t1.Concat(new T[] { t2 }));
        }

        static List<int> Get2IndependentCombinations(IEnumerable<IEnumerable<string>> rounds, List<IEnumerable<int>> roundCombis)
        {
            HashSet<String> independentGroups = new HashSet<string>();
            for (int i = 0; i < roundCombis.Count(); i++)
            {
                var comb = roundCombis[i].ToList();
                independentGroups = new HashSet<string>();
                independentGroups.Add(rounds.ToList()[comb[0]].ToList()[0]);
                independentGroups.Add(rounds.ToList()[comb[0]].ToList()[1]);
                independentGroups.Add(rounds.ToList()[comb[1]].ToList()[0]);
                independentGroups.Add(rounds.ToList()[comb[1]].ToList()[1]);
                if (independentGroups.Count == 4) { return comb; }
            }
            return null;
        }

        static IEnumerable<IEnumerable<T>> GetKCombs<T>(IEnumerable<T> list, int length) where T : IComparable
        {
            if (length == 1) return list.Select(t => new T[] { t });
            return GetKCombs(list, length - 1)
                .SelectMany(t => list.Where(o => o.CompareTo(t.Last()) > 0),
                    (t1, t2) => t1.Concat(new T[] { t2 }));
        }
    }
}
