using Microsoft.Win32;
using System;
using System.Collections.Generic;
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
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template4435
{
    /// <summary>
    /// Логика взаимодействия для TaskActivity.xaml
    /// </summary>
    public partial class TaskActivity : Window
    {
        public TaskActivity()
        {
            InitializeComponent();
        }
        
        

        private void exitBTN_Click(object sender, RoutedEventArgs e)
        {
            this.Hide();
            MainWindow mainWindow = new MainWindow();
            mainWindow.Show();
        }

        private void importBTN_Click(object sender, RoutedEventArgs e)
        {
            OpenFileDialog ofd = new OpenFileDialog()
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "файл Excel (Spisok.xlsx)|*.xlsx",
                Title = "Выберите файл базы данных"
            };
            if (!(ofd.ShowDialog() == true))
                return;


            Excel_Entity data_str = GetData_ToString_FromXL(ofd.FileName);

            var msg = "";
            for(int i = 0; i < data_str.rows; i++)
            {
                for (int j = 0; j < data_str.columns; j++)
                {
                    msg += data_str.data[i, j]+" \t ";
                }
                msg += "\n";

            }
            MessageBox.Show(msg);
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

        public class Excel_Entity {
            public int rows { get; set; }
            public int columns { get; set; }
            public string[,] data { get; set; }
           
        }
    }
}
