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
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace Template4435
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void BnTask_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Сабиров Зульфат Зуфарович","4435_Сабиров_Зульфат");
        }
        private void toWindowBtn_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Мартынов Максим Дмитриевич, 19 лет, группа_4435","4435_Мартынов");
        }
        private void AzatBtn_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Хакимзянов Азат Гайсович", "4435_Хакимзянов_Азат");
        }
        private void BnnTask_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Назмутдинов Рузаль Ильгизович", "4435_Назмутдинов_Рузаль");
        }

        private void LR1_Shumilkin_Click(object sender, RoutedEventArgs e)
        {
            MessageBox.Show("Автор: Шумилкин Александр Олегович", "4435_Шумилкин_Александр");
        }
    }
}
