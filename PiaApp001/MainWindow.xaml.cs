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

namespace PiaApp001
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

        private void btButton1_Click(object sender, RoutedEventArgs e)
        {
            ExcelFunctions excelFunctions = new ExcelFunctions();
            excelFunctions.CreateExcel300Udd(@"c:\temp\TEC_Uddannelser.xlsx", @"c:\temp\TEC_300udd");
        }

        private void btButton2_Click(object sender, RoutedEventArgs e)
        {
            ExcelFunctions excelFunctions = new ExcelFunctions();
            excelFunctions.CreateExcel300Navne(@"c:\temp\TEC_Uddannelser.xlsx", @"c:\temp\TEC_300udd");
        }

        private void btButton3_Click(object sender, RoutedEventArgs e)
        {

        }
    }
}
