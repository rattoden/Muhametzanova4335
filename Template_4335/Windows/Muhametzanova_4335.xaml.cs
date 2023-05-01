using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
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
using Template_4335.Windows.MuhametzanovaAR;

namespace Template_4335.Windows
{
    /// <summary>
    /// Логика взаимодействия для Muhametzanova_4335.xaml
    /// </summary>
    public partial class Muhametzanova_4335 : Window
    {
        public Muhametzanova_4335()
        {
            InitializeComponent();
        }

        private void ExcelPageBtn_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.Navigate(new ExcelPage());
        }

        private void WordPageBtn_Click(object sender, RoutedEventArgs e)
        {
            MainFrame.Navigate(new WordPage());
        }

        private void DeleteDataBtn_Click(object sender, RoutedEventArgs e)
        {
            if (MessageBox.Show("Очистить данные?", "Внимание!", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
            {
                using (var excelEntities = new ExcelEntities())
                {
                    excelEntities.Uslugi.RemoveRange(excelEntities.Uslugi.ToList());
                    excelEntities.SaveChanges();
                    ExcelEntities.GetContext().Uslugi.AsEnumerable().OrderBy(x => Convert.ToInt32(x.Id)).ToList().Clear();
                    foreach (var uslugi in excelEntities.Uslugi.AsEnumerable().OrderBy(x => Convert.ToInt32(x.Id)).ToList())
                    {
                        ExcelEntities.GetContext().Uslugi.AsEnumerable().OrderBy(x => Convert.ToInt32(x.Id)).ToList().Add(uslugi);
                    }
                }
            }
        }
    }
}
