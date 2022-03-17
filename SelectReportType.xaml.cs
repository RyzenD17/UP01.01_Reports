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

namespace UP01._01
{
    /// <summary>
    /// Логика взаимодействия для SelectReportType.xaml
    /// </summary>
    public partial class SelectReportType : Window
    {
        public SelectReportType()
        {
            InitializeComponent();
            CBReportType.Items.Add("Краткая характеристика мо материалам");
            CBReportType.Items.Add("Материала количество на складе которых меньше минимального");
            CBReportType.Items.Add("Все материалы типа Гранулы");
            CBReportType.Items.Add("Метериалы количетво на складе которых равно 300% от минимального");
            CBReportType.Items.Add("Поставщики по их типу");
            CBReportType.SelectedIndex = 0;

        }

        public int Type
        {
            get
            {
                return CBReportType.SelectedIndex + 1;
            }
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }
    }
}
