using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.ComponentModel;
using System.Windows;
using System.Runtime.CompilerServices;
using System.Reflection;
using Microsoft.Win32;
using System.Windows.Controls;
using System.Data;
using System.Data.OleDb;
using System.Xml;
using System.Xml.Serialization;
using Excel = Microsoft.Office.Interop.Excel;

namespace TaskNipom
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private List<Component> electrocomponents = new List<Component>();
        public MainWindow()
        {
            InitializeComponent();
        }

        [Serializable]
        public class Component : INotifyPropertyChanged
        {
            public event PropertyChangedEventHandler PropertyChanged;
            public string nаimenovаnie { get; set; }
            public string proizvoditel { get; set; }
            public string kаtegoriya__montаjа { get; set; }
            public double stoimost { get; set; }
            public double kol_vo { get; set; }
            public double Summa { 
                get { return kol_vo * stoimost; }
                set { }
            }
            private void OnPropertyChanged([CallerMemberName] string propertyName = "")
            {
                if (PropertyChanged != null) 
                {
                    PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
                }
            }
        }
        private void openExcel_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = "*.xls;*.xlsx";
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Выберите документ для загрузки";

            if (openFileDialog.ShowDialog() == true)
            {
                string fileName = openFileDialog.FileName;
                try
                {
                    Excel.Application xlApp = new Excel.Application();
                    Excel.Workbook xlWorkBook = xlApp.Workbooks.Open(fileName);
                    Excel._Worksheet xlWorksheet = xlWorkBook.Sheets[1];
                    Excel.Range xlRange = xlWorksheet.UsedRange;

                    int rowCount = xlRange.Rows.Count;
                    int colCount = xlRange.Columns.Count;

                    int rowCnt = 1;

                    DataSet ds = new DataSet("Electrocomponents");
                    DataTable dt = new DataTable("Component");

                    //Чтение шапки
                    for (int column = 1; column <= colCount; column++)
                    {
                        string value = "";

                        if(xlRange.Cells[rowCnt, column + 1].Value2 == null)
                        {
                            rowCnt = 2; //переходим на следующую строку если первая строка не шапка таблицы
                        }
                        value = (xlRange.Cells[rowCnt, column]).Value2;
                        dt.Columns.Add(value, typeof(string));
                    }

                    //Чтение строк
                    for (rowCnt = rowCnt + 1; rowCnt <= rowCount; rowCnt++)
                    {
                        DataRow dr = dt.NewRow();
                        for (int colCnt = 1; colCnt <= colCount; colCnt++)
                        {
                            if((xlRange.Cells[rowCnt, colCnt]).Value2 != null)
                            {
                                dr[colCnt - 1] = (xlRange.Cells[rowCnt, colCnt]).Value2.ToString();
                            }
                        }
                        dt.Rows.Add(dr);
                    }

                    dt.Columns.Add("Сумма", typeof(string));
                    ds.Tables.Add(dt);
                    DataGrid.ItemsSource = dt.DefaultView;

                    xlApp.Quit();
                }
                catch(Exception ex)
                {
                    MessageBox.Show("Ошибка чтения файла\n" + ex.Message);
                }
            }
        }
        private void openExcelBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.DefaultExt = "*.xls;*.xlsx";
            openFileDialog.Filter = "Excel Files|*.xls;*.xlsx;*.xlsm";
            openFileDialog.Title = "Выберите документ для загрузки";

            if (openFileDialog.ShowDialog() == true)
            {
                string fileName = openFileDialog.FileName;
                string stringCon = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0}; Extended Properties=Excel 12.0;", fileName);

                OleDbConnection dbConnection = new OleDbConnection(stringCon);
                dbConnection.Open();
                
                DataSet ds = new DataSet();

                DataTable schemaTable = dbConnection.GetOleDbSchemaTable(OleDbSchemaGuid.Tables, new object[] { null, null, null, "TABLE" });

                string sheet1 = (string)schemaTable.Rows[0].ItemArray[2];
                string select = string.Format("SELECT * FROM [{0}]", sheet1);

                OleDbDataAdapter adapter = new OleDbDataAdapter(select, dbConnection);

                adapter.Fill(ds);

                DataTable tb = ds.Tables[0];

                //костыль для файла "Исходные данные.xlsx"
                tb.Rows.RemoveAt(0);
  
                this.DataTable_to_Class(tb);
                DataGrid.ItemsSource = electrocomponents;
                dbConnection.Close();
            }
        }
        private void opentXmlBtn_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "XML Files|*.xml";
            openFileDialog.Title = "Выберите документ для загрузки";

            if (openFileDialog.ShowDialog() == true)
            {
                try
                {
                    Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
                    XmlSerializer xml = new XmlSerializer(typeof(List<Component>), new XmlRootAttribute("ElectroComponents"));

                    using (FileStream fs = new FileStream(openFileDialog.FileName, FileMode.OpenOrCreate))
                    {
                        electrocomponents = (List<Component>)xml.Deserialize(fs);
                    }
                }
                catch(XmlException ex)
                {
                    MessageBox.Show("Ошибка чтения XML файла \n" + ex.Message);
                }
                DataGrid.ItemsSource = electrocomponents;
            }
        }
        private void saveXmlBtn_Click(object sender, RoutedEventArgs e)
        {
            SaveFileDialog saveFileDialog = new SaveFileDialog();
            saveFileDialog.Filter = "XML|*.xml";
            saveFileDialog.Title = "Сохраните XML документ";

            if (saveFileDialog.ShowDialog() == true)
            {
                try
                {
                    XmlSerializer xml = new XmlSerializer(typeof(List<Component>), new XmlRootAttribute("ElectroComponents"));

                    using(FileStream fs = new FileStream(saveFileDialog.FileName, FileMode.OpenOrCreate))
                    {
                        xml.Serialize(fs, electrocomponents);
                    }
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Ошибка записи XML файла \n" + ex.Message);
                }
            }
        }
        private void DataTable_to_Class(DataTable dataTable)
        {
            foreach (DataRow row in dataTable.Rows)
            {
                electrocomponents.Add(
                        new Component()
                        {
                            nаimenovаnie = row.ItemArray[0].ToString(),
                            proizvoditel = row.ItemArray[1].ToString(),
                            kаtegoriya__montаjа = row.ItemArray[2].ToString(),
                            stoimost = Convert.ToDouble(row.ItemArray[3]),
                            kol_vo = Convert.ToDouble(row.ItemArray[4]),
                            //Summa = Convert.ToDouble(row.ItemArray[3]) * Convert.ToDouble(row.ItemArray[4])
                        });
            }
        }
        private void DataGrid_CellEditEnding(object sender, DataGridCellEditEndingEventArgs e)
        {
            Component el = e.Row.Item as Component;
            if (el != null)
            {
                el.Summa = el.kol_vo * el.stoimost;
            }

            DataGrid.ItemsSource = null;
            DataGrid.ItemsSource = electrocomponents;

            MessageBox.Show($"Новая сумма {el.Summa}");
        }
    }
}
