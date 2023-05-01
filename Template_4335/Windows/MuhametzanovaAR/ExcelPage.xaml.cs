using Microsoft.Office.Interop.Excel;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;

namespace Template_4335.Windows.MuhametzanovaAR
{
    /// <summary>
    /// Логика взаимодействия для ExcelPage.xaml
    /// </summary>
    public partial class ExcelPage : System.Windows.Controls.Page
    {
        public ExcelPage()
        {
            InitializeComponent();
            DBGridModel.ItemsSource = ExcelEntities.GetContext().Uslugi.AsEnumerable().OrderBy(x => Convert.ToInt32(x.Id)).ToList();
        }

        private void ImportBtn_Click(object sender, RoutedEventArgs e)
        {
            var openFileDialog = new OpenFileDialog
            {
                DefaultExt = "*.xls;*.xlsx",
                Filter = "Файл Excel|*.xlsx",
                Title = "Выберите файл"
            };
            if (!openFileDialog.ShowDialog() == true)
                return;
            ImportData(openFileDialog.FileName);
        }
        private static void ImportData(string path)
        {
            try
            {
                var objWorkExcel = new Excel.Application();
                var objWorkBook = objWorkExcel.Workbooks.Open(path);
                var objWorkSheet = (Excel.Worksheet)objWorkBook.Sheets[1];
                var lastCell = objWorkSheet.Cells.SpecialCells(Excel.XlCellType.xlCellTypeLastCell);
                var columns = lastCell.Column;
                var rows = lastCell.Row;
                var list = new string[rows, columns];
                for (var j = 0; j < columns; j++)
                {
                    for (var i = 1; i < rows; i++)
                    {
                        list[i, j] = objWorkSheet.Cells[i + 1, j + 1].Text;
                    }
                }

                objWorkBook.Close(false, Type.Missing, Type.Missing);
                objWorkExcel.Quit();
                GC.Collect();
                using (var db = new ExcelEntities())
                {
                    for (var i = 1; i < 51; i++)
                    {
                        var uslugi = new Uslugi
                        {
                            Id = list[i, 0].ToString(),
                            IdZakaza = list[i, 1].ToString(),
                            DataSozdaniya = list[i, 2].ToString(),
                            VremyaZakaza = list[i, 3].ToString(),
                            IdClienta = list[i, 4].ToString(),
                            Uslugii = list[i, 5].ToString(),
                            Statuss = list[i, 6].ToString(),
                            DataZakritiya = list[i, 7].ToString(),
                            VremyaProkata = int.Parse(list[i, 8])
                        };
                        db.Uslugi.Add(uslugi);
                    }
                    try
                    {
                        db.SaveChanges();
                        MessageBox.Show("Данные импортированы!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show(ex.Message, "Внимание", MessageBoxButton.OK);
                    }

                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message, "Внимание1", MessageBoxButton.OK);
            }
        }

        private void ExportBtn_Click(object sender, RoutedEventArgs e)
        {
            #region Объявление листов
            var first = new List<Uslugi>();
            var second = new List<Uslugi>();
            var third = new List<Uslugi>();
            var fourth = new List<Uslugi>();
            var fifth = new List<Uslugi>();
            var sixth = new List<Uslugi>();
            var seventh = new List<Uslugi>();
            var categoriesPriceCount = 7;
            #endregion

            using (var excelEntities = new ExcelEntities())
            {
                #region Сортировка
                first = excelEntities.Uslugi.ToList().Where(sR => Convert.ToInt32(sR.VremyaProkata) == 120).OrderBy(fR => fR.IdZakaza).ToList();
                second = excelEntities.Uslugi.ToList().Where(sR => Convert.ToInt32(sR.VremyaProkata) == 240).OrderBy(fR => fR.IdZakaza).ToList();
                third = excelEntities.Uslugi.ToList().Where(sR => Convert.ToInt32(sR.VremyaProkata) == 320).OrderBy(fR => fR.IdZakaza).ToList();
                fourth = excelEntities.Uslugi.ToList().Where(sR => Convert.ToInt32(sR.VremyaProkata) == 360).OrderBy(fR => fR.IdZakaza).ToList();
                fifth = excelEntities.Uslugi.ToList().Where(sR => Convert.ToInt32(sR.VremyaProkata) == 480).OrderBy(fR => fR.IdZakaza).ToList();
                sixth = excelEntities.Uslugi.ToList().Where(sR => Convert.ToInt32(sR.VremyaProkata) == 600).OrderBy(fR => fR.IdZakaza).ToList();
                seventh = excelEntities.Uslugi.ToList().Where(sR => Convert.ToInt32(sR.VremyaProkata) == 720).OrderBy(fR => fR.IdZakaza).ToList();
                #endregion

                var app = new Excel.Application { SheetsInNewWorkbook = categoriesPriceCount };
                var book = app.Workbooks.Add(Type.Missing);

                #region Создание листов в Excel
                var startRowIndex = 1;
                var sheet1 = app.Worksheets.Item[1];
                sheet1.Name = "Время проката 2 часа";
                var sheet2 = app.Worksheets.Item[2];
                sheet2.Name = "Время проката 4 часа";
                var sheet3 = app.Worksheets.Item[3];
                sheet3.Name = "Время проката 5,3 часа";
                var sheet4 = app.Worksheets.Item[4];
                sheet4.Name = "Время проката 6 часов";
                var sheet5 = app.Worksheets.Item[5];
                sheet5.Name = "Время проката 8 часов";
                var sheet6 = app.Worksheets.Item[6];
                sheet6.Name = "Время проката 10 часов";
                var sheet7 = app.Worksheets.Item[7];
                sheet7.Name = "Время проката 12 часов";
                #endregion

                #region Создание колонок в Excel
                sheet1.Cells[1][startRowIndex] = "Id";
                sheet1.Cells[2][startRowIndex] = "Код заказа";
                sheet1.Cells[3][startRowIndex] = "Дата создания";
                sheet1.Cells[4][startRowIndex] = "Код клиента";
                sheet1.Cells[5][startRowIndex] = "Услуги";

                sheet2.Cells[1][startRowIndex] = "Id";
                sheet2.Cells[2][startRowIndex] = "Код заказа";
                sheet2.Cells[3][startRowIndex] = "Дата создания";
                sheet2.Cells[4][startRowIndex] = "Код клиента";
                sheet2.Cells[5][startRowIndex] = "Услуги";

                sheet3.Cells[1][startRowIndex] = "Id";
                sheet3.Cells[2][startRowIndex] = "Код заказа";
                sheet3.Cells[3][startRowIndex] = "Дата создания";
                sheet3.Cells[4][startRowIndex] = "Код клиента";
                sheet3.Cells[5][startRowIndex] = "Услуги";

                sheet4.Cells[1][startRowIndex] = "Id";
                sheet4.Cells[2][startRowIndex] = "Код заказа";
                sheet4.Cells[3][startRowIndex] = "Дата создания";
                sheet4.Cells[4][startRowIndex] = "Код клиента";
                sheet4.Cells[5][startRowIndex] = "Услуги";

                sheet5.Cells[1][startRowIndex] = "Id";
                sheet5.Cells[2][startRowIndex] = "Код заказа";
                sheet5.Cells[3][startRowIndex] = "Дата создания";
                sheet5.Cells[4][startRowIndex] = "Код клиента";
                sheet5.Cells[5][startRowIndex] = "Услуги";

                sheet6.Cells[1][startRowIndex] = "Id";
                sheet6.Cells[2][startRowIndex] = "Код заказа";
                sheet6.Cells[3][startRowIndex] = "Дата создания";
                sheet6.Cells[4][startRowIndex] = "Код клиента";
                sheet6.Cells[5][startRowIndex] = "Услуги";

                sheet7.Cells[1][startRowIndex] = "Id";
                sheet7.Cells[2][startRowIndex] = "Код заказа";
                sheet7.Cells[3][startRowIndex] = "Дата создания";
                sheet7.Cells[4][startRowIndex] = "Код клиента";
                sheet7.Cells[5][startRowIndex] = "Услуги";
                startRowIndex++;
                #endregion

                #region Заполнение первого листа
                for (var i = 0; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in first)
                    {
                        sheet1.Cells[1][startRowIndex] = item.Id;
                        sheet1.Cells[2][startRowIndex] = item.IdZakaza;
                        sheet1.Cells[3][startRowIndex] = item.DataSozdaniya.ToString();
                        sheet1.Cells[4][startRowIndex] = item.IdClienta;
                        sheet1.Cells[5][startRowIndex] = item.Uslugii;
                        startRowIndex++;
                    }
                }
                #endregion

                #region Заполнение второго листа
                for (var i = 1; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in second)
                    {
                        sheet2.Cells[1][startRowIndex] = item.Id;
                        sheet2.Cells[2][startRowIndex] = item.IdZakaza;
                        sheet2.Cells[3][startRowIndex] = item.DataSozdaniya.ToString();
                        sheet2.Cells[4][startRowIndex] = item.IdClienta;
                        sheet2.Cells[5][startRowIndex] = item.Uslugii;
                        startRowIndex++;
                    }
                }
                #endregion

                #region Заполнение третьего листа
                for (var i = 2; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in third)
                    {
                        sheet3.Cells[1][startRowIndex] = item.Id;
                        sheet3.Cells[2][startRowIndex] = item.IdZakaza;
                        sheet3.Cells[3][startRowIndex] = item.DataSozdaniya.ToString();
                        sheet3.Cells[4][startRowIndex] = item.IdClienta;
                        sheet3.Cells[5][startRowIndex] = item.Uslugii;
                        startRowIndex++;
                    }
                }
                #endregion

                #region Заполнение четвёртого листа
                for (var i = 3; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in fourth)
                    {
                        sheet4.Cells[1][startRowIndex] = item.Id;
                        sheet4.Cells[2][startRowIndex] = item.IdZakaza;
                        sheet4.Cells[3][startRowIndex] = item.DataSozdaniya.ToString();
                        sheet4.Cells[4][startRowIndex] = item.IdClienta;
                        sheet4.Cells[5][startRowIndex] = item.Uslugii;
                        startRowIndex++;
                    }
                }
                #endregion

                #region Заполнение пятого листа
                for (var i = 4; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in fifth)
                    {
                        sheet5.Cells[1][startRowIndex] = item.Id;
                        sheet5.Cells[2][startRowIndex] = item.IdZakaza;
                        sheet5.Cells[3][startRowIndex] = item.DataSozdaniya.ToString();
                        sheet5.Cells[4][startRowIndex] = item.IdClienta;
                        sheet5.Cells[5][startRowIndex] = item.Uslugii;
                        startRowIndex++;
                    }
                }
                #endregion

                #region Заполнение шестого листа
                for (var i = 5; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in sixth)
                    {
                        sheet6.Cells[1][startRowIndex] = item.Id;
                        sheet6.Cells[2][startRowIndex] = item.IdZakaza;
                        sheet6.Cells[3][startRowIndex] = item.DataSozdaniya.ToString();
                        sheet6.Cells[4][startRowIndex] = item.IdClienta;
                        sheet6.Cells[5][startRowIndex] = item.Uslugii;
                        startRowIndex++;
                    }
                }
                #endregion

                #region Заполнение седьмого листа
                for (var i = 6; i < categoriesPriceCount; i++)
                {
                    startRowIndex = 2;
                    foreach (var item in seventh)
                    {
                        sheet7.Cells[1][startRowIndex] = item.Id;
                        sheet7.Cells[2][startRowIndex] = item.IdZakaza;
                        sheet7.Cells[3][startRowIndex] = item.DataSozdaniya.ToString();
                        sheet7.Cells[4][startRowIndex] = item.IdClienta;
                        sheet7.Cells[5][startRowIndex] = item.Uslugii;
                        startRowIndex++;
                    }
                }
                #endregion

                app.Visible = true;
            }
        }
    }
}
