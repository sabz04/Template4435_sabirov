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
using Word = Microsoft.Office.Interop.Word;

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
            var list_times = excel_data.Select(x => DateTime.Parse(x.CreateDate.ToString()).ToShortDateString()).Distinct().OrderBy(x=>x).ToList();

            var app = new Word.Application();

            
            Word.Document document = app.Documents.Add();
            int counter = 0;
            for (int j=0; j< list_times.Count(); j++)
            {
                Word.Paragraph paragraph = document.Paragraphs.Add();
                Word.Range range = paragraph.Range;
                range.Text = list_times[j].ToString();
                paragraph.set_Style("Заголовок 1");
                range.InsertParagraphAfter();

                Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Word.Range tableRange = tableParagraph.Range;

                Word.Table excel_dataTable =
document.Tables.Add(tableRange, excel_data.Select(x => DateTime.Parse(x.CreateDate).ToShortDateString().ToString()).Where(x => range.Text.Contains(x)).ToList().Count() + 1, 4);
                int i = 1;
                foreach (var item in excel_data)
                {
                    string list_Dt_str = DateTime.Parse(item.CreateDate).ToShortDateString().ToString();
                    if (range.Text.Contains(list_Dt_str))
                    {
                        
                        excel_dataTable.Borders.InsideLineStyle =
                        excel_dataTable.Borders.OutsideLineStyle =
                        Word.WdLineStyle.wdLineStyleSingle;
                        excel_dataTable.Range.Cells.VerticalAlignment =
                        Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;
                        Word.Range cellRange = excel_dataTable.Cell(1, 1).Range;

                        cellRange.Text = "ID";
                        cellRange = excel_dataTable.Cell(1, 2).Range;
                        cellRange.Text = "Код заказа";
                        cellRange = excel_dataTable.Cell(1, 3).Range;
                        cellRange.Text = "Код клиента";
                        cellRange = excel_dataTable.Cell(1, 4).Range;
                        cellRange.Text = "Услуги";

                        excel_dataTable.Rows[1].Range.Bold = 1;
                        excel_dataTable.Rows[1].Range.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        cellRange = excel_dataTable.Cell(i + 1, 1).Range;
                        cellRange.Text = item.Id.ToString();
                        cellRange.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        cellRange = excel_dataTable.Cell(i + 1, 2).Range;
                        cellRange.Text = item.CodeOrder;
                        cellRange.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        cellRange = excel_dataTable.Cell(i + 1, 3).Range;
                        cellRange.Text = item.CodeClient;
                        cellRange.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;

                        cellRange = excel_dataTable.Cell(i + 1, 4).Range;
                        cellRange.Text = item.Services;
                        cellRange.ParagraphFormat.Alignment =
                        Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        i++;
                    }
                }
            }


            app.Visible = true;
        }
    }
}


