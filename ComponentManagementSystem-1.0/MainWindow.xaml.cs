using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Controls.Primitives;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Newtonsoft.Json;
namespace ComponentManagementSystem_1._0
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        // Taking all the class variable
        private string _ContextName = "ComponentManagementSystem";
        private string _ProcessFileName;
        private string _DataFileName;
        public int _CurrentRowIndex = -1;
        public string _CurrentRowComponentValue = "";
        public int _ProcessStartIndex = 0;
        public List<String> _Processes;
        public List<string> _Columns;
        public List<List<String>> _Rows;
        public DataTable _DataTable;
        public DataRow _DataRow;
        public DataColumn _DataColumn;
        public Dictionary<string, int> _ProcessMap;
        public double _MaxWidthOfProcesses = 0;
        public int _TotalNumberOfProcesses = 30;
        public string _FilePath;
        public string _DirectoryPath;
        public MainWindow()
        {
            InitializeComponent();
            // This is for making a folder named as project name in current user folder
            // In my case it is under vinay folder
            // This is used for making a process file under this folder.
            _DirectoryPath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\" + _ContextName;
            _FilePath = Environment.GetFolderPath(Environment.SpecialFolder.UserProfile) + "\\" + _ContextName + "\\data.json";
            CheckProcessFileExists();
        }


        private void btnProcess_Click(object sender, RoutedEventArgs e)
        {
            // This function will invoke when we click on any process button.
            if (!CheckValidRow(_CurrentRowIndex)) return;
            // We check for next empty cell from this current row by using process map
            // and then put data in that cell.
            Button b = sender as Button;
            DataRowView dataRow = mainDataGrid.Items.GetItemAt(_CurrentRowIndex) as DataRowView;
            dataRow.BeginEdit();
            dataRow[_ProcessMap[_CurrentRowComponentValue]] = b.Content.ToString();
            _ProcessMap[_CurrentRowComponentValue]++;
            dataRow.EndEdit();

            mainDataGrid.CurrentCell = new DataGridCellInfo(mainDataGrid.Items[_CurrentRowIndex], mainDataGrid.Columns[_ProcessMap[_CurrentRowComponentValue]]);
            mainDataGrid.BeginEdit();
            //updateDataTable(_CurrentRowIndex, b.Content.ToString());

        }
        private Boolean CheckValidRow(int currentRow)
        {
            // This function check that, is any cell from dataGrid is selected or not.
            // As firstly we have to select a cell from data grid, then current row value will change from -1 to some other value
            if (currentRow == -1 || currentRow > _Processes.Count())
            {
                txtError.Text = "Error : No Row Selected";
                return false;
            }
            return true;
        }
        private void btnGenerateExcel_Click(object sender, RoutedEventArgs e)
        {
            // This function will generate excel file, by taking data from data grid.
            // firstly it will take saveFileName from user
            string saveFileName = "";
            Microsoft.Win32.SaveFileDialog dlg = new Microsoft.Win32.SaveFileDialog();
            dlg.FileName = "Report"; // Default file name
            dlg.DefaultExt = ".xlsx"; // Default file extension
            dlg.Filter = "Excel Document (.xlsx)|*.xlsx"; // Filter files by extension

            // Show save file dialog box
            Nullable<bool> result = dlg.ShowDialog();

            // Process save file dialog box results
            if (result == true)
            {
                // Save document
                saveFileName = dlg.FileName;
            }
            XSSFWorkbook workbook = new XSSFWorkbook();
            XSSFFont myFont = (XSSFFont)workbook.CreateFont();
            myFont.FontHeightInPoints = 11;
            myFont.FontName = "Tahoma";


            // Defining a border
            XSSFCellStyle numberCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            numberCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Right;

            XSSFCellStyle CharactersCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            CharactersCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Left;

            XSSFCellStyle ColumnCellStyle = (XSSFCellStyle)workbook.CreateCellStyle();
            ColumnCellStyle.Alignment = NPOI.SS.UserModel.HorizontalAlignment.Center;

            ISheet Sheet = workbook.CreateSheet("Report");

            int rowIndex = 0;
            IRow HeaderRow = Sheet.CreateRow(rowIndex++);
            for (int i = 0; i < _DataTable.Columns.Count; i++)
            {
                CreateCell(HeaderRow, i, _DataTable.Columns[i].ColumnName, ColumnCellStyle);
            }

            for (int i = 0; i < _DataTable.Rows.Count; i++)
            {
                IRow currentRow = Sheet.CreateRow(rowIndex++);
                for (int j = 0; j < _DataTable.Columns.Count; j++)
                {
                    if (_DataTable.Rows[i][j].GetType() != DBNull.Value.GetType())
                    {
                        if (_DataTable.Rows[i][j].ToString().All(char.IsDigit)) CreateCell(currentRow, j, _DataTable.Rows[i][j].ToString(), numberCellStyle);
                        else CreateCell(currentRow, j, _DataTable.Rows[i][j].ToString(), CharactersCellStyle);
                    }
                }
            }

            int lastColumNum = Sheet.GetRow(0).LastCellNum;
            for (int i = 0; i <= lastColumNum; i++)
            {
                if (i == 0 || i == 1 || i == 3) Sheet.SetColumnWidth(i, (3 + 1) * 256);
                else Sheet.SetColumnWidth(i, 2300);
                //Sheet.AutoSizeColumn(i);
                GC.Collect();
            }
            // Write Excel to disk 
            using (var fileData = new FileStream(saveFileName, FileMode.Create))
            {
                workbook.Write(fileData);
            }

        }
        private void CreateCell(IRow CurrentRow, int CellIndex, string Value, XSSFCellStyle Style)
        {
            // This function will create cell in out excel file
            ICell Cell = CurrentRow.CreateCell(CellIndex);
            if (Value.All(char.IsDigit)) Cell.SetCellValue(int.Parse(Value));
            else Cell.SetCellValue(Value);
            Cell.CellStyle = Style;
        }

        private void dataGrid_dblClick(object sender, MouseButtonEventArgs e)
        {
            // This function will take value of row where we double click
            try
            {
                _CurrentRowIndex = mainDataGrid.Items.IndexOf(mainDataGrid.CurrentItem);
                DataRowView dataRow = mainDataGrid.Items.GetItemAt(_CurrentRowIndex) as DataRowView;
                string cellValue = dataRow.Row.ItemArray[2].ToString();
                _CurrentRowComponentValue = cellValue;
                // take value of next empty cell and blink cursor there
                mainDataGrid.CurrentCell = new DataGridCellInfo(mainDataGrid.Items[_CurrentRowIndex], mainDataGrid.Columns[_ProcessMap[_CurrentRowComponentValue]]);
                mainDataGrid.BeginEdit();
            }
            catch (Exception ex)
            {
                txtError.Text = "Error : " + ex.Message;
            }
        }
        private void CheckProcessFileExists()
        {
            // Checking if directory exists, if not, we create a directory
            if (!Directory.Exists(_DirectoryPath))
                System.Diagnostics.Trace.WriteLine(Directory.CreateDirectory(_DirectoryPath));
            // if the process file not exists in current directory we will return from this function
            if (!File.Exists(_FilePath)) return;
            // if process file exists, we will take file name and then draw buttons accordingly.
            string jsonString = File.ReadAllText(_FilePath);
            FileStore fs = JsonConvert.DeserializeObject<FileStore>(jsonString);
            if (fs._ProcessFile != null && File.Exists(fs._ProcessFile))
            {
                _ProcessFileName = fs._ProcessFile;
                PopulateProcessButtons();
            }
        }
        private void btnUploadProcesses_Click(object sender, RoutedEventArgs e)
        {
            // this function is used when process file not exists, and it will invoke in upload process button click
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xlsx"; // Default file extension
            dlg.Filter = "Excel Document (.xlsx)|*.xlsx"; // Filter files by extension

            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                _ProcessFileName = dlg.FileName;
            }
            // This object is used for storing process file.
            FileStore fileStore = new FileStore();
            fileStore._ProcessFile = _ProcessFileName;
            var jsonString = JsonConvert.SerializeObject(fileStore);
            File.WriteAllText(_FilePath, jsonString);
            PopulateProcessButtons();
        }
        public void PopulateProcessButtons()
        {
            // This function will read data from process file and make a list from that data
            IWorkbook workbook = null;
            if (!File.Exists(_ProcessFileName))
            {
                txtError.Text = "Error : File not exists";
                return;
            }
            FileStream fileStream = new FileStream(_ProcessFileName, FileMode.Open, FileAccess.Read);
            if (_ProcessFileName.IndexOf(".xlsx") > 0) workbook = new XSSFWorkbook(fileStream);
            else if (_ProcessFileName.IndexOf(".xls") > 0) workbook = new HSSFWorkbook(fileStream);
            ISheet sheet = workbook.GetSheetAt(0);
            if (sheet != null)
            {
                int rowCount = sheet.LastRowNum;
                _Processes = new List<string>();
                for (int i = 1; i <= rowCount; i++)
                {
                    IRow currRow = sheet.GetRow(i);
                    string processName = currRow.GetCell(1).StringCellValue;
                    _Processes.Add(processName);
                }
                MakeProcessPanel();
            }
        }
        public void MakeProcessPanel()
        {
            // this function will make buttons from taking data from processList.
            if (_Processes.Count == 0)
            {
                txtError.Text = "Error : You did not feed process file. Here process count is zero";
            }
            _Processes.Sort();
            StackPanel stackPanel = null;
            for (int i = 0, columnIndex = 0; i < _Processes.Count; i++)
            {
                if (i % 25 == 0)
                {
                    stackPanel = new StackPanel();
                    Grid.SetColumn(stackPanel, columnIndex++);
                    processGrid.Children.Add(stackPanel);
                }
                Button b = new Button()
                {
                    Content = _Processes[i],
                    Style = FindResource("ProcessBtnStyle") as Style
                };
                b.Click += new RoutedEventHandler(btnProcess_Click);
                var size = new Size(double.PositiveInfinity, double.PositiveInfinity);
                b.Measure(size);
                b.Arrange(new Rect(b.DesiredSize));
                if (b.ActualWidth > _MaxWidthOfProcesses) _MaxWidthOfProcesses = b.ActualWidth;
                stackPanel.Children.Add(b);
            }
        }
        private void btnUploadComponents_Click(object sender, RoutedEventArgs e)
        {
            // This function take a component file 
            Microsoft.Win32.OpenFileDialog dlg = new Microsoft.Win32.OpenFileDialog();
            dlg.DefaultExt = ".xlsx"; // Default file extension
            dlg.Filter = "Excel Document (.xlsx)|*.xlsx"; // Filter files by extension

            Nullable<bool> result = dlg.ShowDialog();

            if (result == true)
            {
                _DataFileName = dlg.FileName;
            }


            PopulateDataGrid();
        }
        private void PopulateDataGrid()
        {
            // this process make list of _Rows by taking data from compoent excel file
            IWorkbook workbook2 = null;
            FileStream fileStream2 = new FileStream(_DataFileName, FileMode.Open, FileAccess.Read);
            if (_DataFileName.IndexOf(".xlsx") > 0) workbook2 = new XSSFWorkbook(fileStream2);
            else if (_DataFileName.IndexOf(".xls") > 0) workbook2 = new HSSFWorkbook(fileStream2);
            ISheet sheet2 = workbook2.GetSheetAt(0);
            if (sheet2 != null)
            {
                IRow firstRow = sheet2.GetRow(1);
                _Columns = new List<string>();
                _Rows = new List<List<String>>();
                _Columns.Add(firstRow.GetCell(0).StringCellValue);
                _Columns.Add(firstRow.GetCell(1).StringCellValue);
                _Columns.Add(firstRow.GetCell(2).StringCellValue);
                _Columns.Add(firstRow.GetCell(3).StringCellValue);
                _ProcessStartIndex = 4;
                for (int i = 1; i < _TotalNumberOfProcesses; i++)
                {
                    _Columns.Add("P" + i);
                }
                int rowCount = sheet2.LastRowNum;
                for (int i = 2; i <= rowCount; i++)
                {
                    IRow currRow = sheet2.GetRow(i);
                    List<string> row = new List<String>();
                    row.Add(currRow.GetCell(0).NumericCellValue.ToString());
                    row.Add(currRow.GetCell(1).NumericCellValue.ToString());
                    row.Add(currRow.GetCell(2).StringCellValue);
                    row.Add(currRow.GetCell(3).NumericCellValue.ToString());
                    _Rows.Add(row);
                }
                //mainDataGrid.Loaded += SetMinWidths;
                MakeDataTable();
                SetWidthOfElements();
            }
        }
        public void MakeDataTable()
        {
            // This function will take data from _Rows list and make data grid, by using data table
            _ProcessMap = new Dictionary<string, int>();
            _DataTable = new DataTable();
            for (int i = 0; i < _Columns.Count; i++)
            {
                _DataColumn = new DataColumn(_Columns[i]);
                _DataTable.Columns.Add(_DataColumn);
            }
            for (int i = 0; i < _Rows.Count; i++)
            {
                _DataRow = _DataTable.NewRow();
                _ProcessMap.Add(_Rows[i][2], 4);
                for (int j = 0; j < _Rows[i].Count; j++)
                {
                    _DataRow[_Columns[j]] = _Rows[i][j];
                }
                _DataTable.Rows.Add(_DataRow);
            }
            mainDataGrid.Style = FindResource("DataGridStyle") as Style;
            mainDataGrid.ItemsSource = _DataTable.DefaultView;
        }

        public void SetWidthOfElements()
        {
            // this function will make width of all the cells in our data grid
            foreach (var column in mainDataGrid.Columns)
            {
                column.MinWidth = column.ActualWidth;
                column.Width = _MaxWidthOfProcesses;
            }
            // This are some _Columns with numbers only, and we want their width to be only upto 3 numbers come in cell, so we set their width manually
            mainDataGrid.Columns[0].Width = 35;
            mainDataGrid.Columns[1].Width = 35;
            mainDataGrid.Columns[3].Width = 35;
        }
        public class FileStore
        {
            /// This class used for storing process file name as a json
            public string _ProcessFile { set; get; }
        }

    }
}
