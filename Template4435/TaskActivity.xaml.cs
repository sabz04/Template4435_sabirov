using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;


namespace Template4435
{
    /// <summary>
    /// Логика взаимодействия для TaskActivity.xaml
    /// </summary>
    public partial class TaskActivity : Window
    {
        public List<ExcelData> excel_data;
        public TaskActivity()
        {
            InitializeComponent();
            Refresh();
        }
        
        private void Refresh()
        {
            using (DataModelContainer db = new DataModelContainer())
            {
                excelGrid.Items.Clear();
                excelGrid.ItemsSource = null;
                excelGrid.ItemsSource = db.ExcelDataSet.ToList();
            }
        }

        private void exitBTN_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
        }

        private void importBTN_Click(object sender, RoutedEventArgs e)
        {
            using (DataModelContainer excelEntity = new DataModelContainer())
            {
                if (excelEntity.ExcelDataSet.Count() > 0)
                {
                    MessageBox.Show("Очистите базу данных для предотвращения дальнейших ошибок.");
                    return;
                }
            }
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx;*.json",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx|Json files (*.json)|*.json|Text files (*.txt)|*.txt",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;

            if (ofd.FileName.Contains("json"))
            {
                MessageBox.Show($"{ofd.FileName}");

                GetDataFrom_Json(ofd.FileName);
                return;
            }


            Excel_Entity data_str = GetData_ToString_FromXL(ofd.FileName);

            using (DataModelContainer excelEntity = new DataModelContainer())
            {
                for (int i = 0; i < data_str.rows; i++)
                {
                    if (data_str.data[i, 1] == "" || data_str.data[i, 1] == " ")
                        continue;
                    excelEntity.ExcelDataSet.Add(new ExcelData()
                    {
                        Id = i,
                        CodeOrder = data_str.data[i, 1],
                        CreateDate = data_str.data[i, 2],
                        CreateTime = data_str.data[i, 3],
                        CodeClient = data_str.data[i, 4],
                        Services = data_str.data[i, 5],
                        Status = data_str.data[i, 6],
                        ClosedDate = data_str.data[i, 7],
                        ProkatTime = data_str.data[i, 8],

                    });
                }
                excelEntity.SaveChanges();
                excelGrid.ItemsSource = excelEntity.ExcelDataSet.ToList();
            }


        }
        private void GetDataFrom_Json(string file)
        {
            List<ExcelData> datas = new List<ExcelData>();
            using (FileStream fs = new FileStream(file, FileMode.OpenOrCreate))
            {
                datas = JsonSerializer.Deserialize<List<ExcelData>>(fs);
                
            }
            using(DataModelContainer db = new DataModelContainer())
            {
                foreach(var item in datas)
                {
                    db.ExcelDataSet.Add(item);
                }
                db.SaveChanges();
                Refresh();
            }

        }
       
        private Excel_Entity GetData_ToString_FromXL(string url)
        {
            string[,] list;

            Excel.Application ObjWorkExcel = new Excel.Application();

            Excel.Workbook ObjWorkBook = ObjWorkExcel.Workbooks.Open(url);

            Excel.Worksheet ObjWorkSheet = (Excel.Worksheet)ObjWorkBook.Sheets[1];

            var lastCell = ObjWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
            int _columns = (int)lastCell.Column;
            int _rows = (int)lastCell.Row;
            list = new string[_rows, _columns];
            for (int j = 0; j < _columns; j++)
                for (int i = 0; i < _rows; i++)
                    list[i, j] = ObjWorkSheet.Cells[i + 1, j + 1].Text;

            Excel_Entity ent =
                new Excel_Entity();

            ent.data = list;
            ent.columns = _columns;
            ent.rows = _rows;

            ObjWorkBook.Close(false, Type.Missing, Type.Missing);
            ObjWorkExcel.Quit();
            GC.Collect();

            return
                ent;
        }

        public class Excel_Entity
        {
            public int rows { get; set; }
            public int columns { get; set; }
            public string[,] data { get; set; }

        }

        private void clearBTN_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                using (DataModelContainer db = new DataModelContainer())
                {
                    foreach (var row in db.ExcelDataSet)
                    {
                        db.ExcelDataSet.Remove(row);

                    }
                    db.SaveChanges();
                    excelGrid.ItemsSource = null;
                    excelGrid.Items.Clear();
                }
                MessageBox.Show("Готово!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Ошибка! Возможно, все уже и так пусто!" + ex.Message);
            }
        }

        private void exportBTN_Click(object sender, RoutedEventArgs e)
        {
            using (var db = new DataModelContainer())
            {
                excel_data = db.ExcelDataSet.ToList();
            }
            var list_times = excel_data.Select(x => x.ProkatTime).Distinct().ToList();
            list_times.RemoveAt(0);


            var app = new Excel.Application();
            app.SheetsInNewWorkbook = list_times.Count();
            Excel.Workbook workbook = app.Workbooks.Add(Type.Missing);

            for (int i = 0; i < list_times.Count(); i++)
            {
                Excel.Worksheet worksheet = app.Worksheets.Item[i + 1];
                worksheet.Name = Convert.ToString(list_times[i]);

                int j = 1;
                worksheet.Cells[1][j] = "ID";
                worksheet.Cells[2][j] = "Код заказа";
                worksheet.Cells[3][j] = "Дата создания";
                worksheet.Cells[4][j] = "Код клиента";
                worksheet.Cells[5][j] = "Услуги";
                j = 2;
                foreach (var item in excel_data)
                {
                    if (item.ProkatTime == worksheet.Name)
                    {
                        worksheet.Cells[1][j] = item.Id;
                        worksheet.Cells[2][j] = item.CodeOrder;
                        worksheet.Cells[3][j] = item.CreateTime;
                        worksheet.Cells[4][j] = item.CodeClient;
                        worksheet.Cells[5][j] = item.Services;
                        j++;
                    }
                }
            }
            app.Visible = true;
        }

        private void exportWordBTN_Click(object sender, RoutedEventArgs e)
        {
            using (var db = new DataModelContainer())
            {
                excel_data = db.ExcelDataSet.ToList();
            }
            var list_times = excel_data.Select(x => DateTime.Parse(x.CreateTime.ToString())).Distinct().OrderBy(x=>x).ToList();
           
        }
    }
}


