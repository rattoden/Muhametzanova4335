using Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
using System;
using System.Collections.Generic;
using System.Data.Entity.Validation;
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
using System.Windows.Navigation;
using System.Windows.Shapes;
using Word = Microsoft.Office.Interop.Word;

namespace Template_4335.Windows.MuhametzanovaAR
{
    /// <summary>
    /// Логика взаимодействия для WordPage.xaml
    /// </summary>
    public partial class WordPage : System.Windows.Controls.Page
    {
        public WordPage()
        {
            InitializeComponent();
            DBGridModel.ItemsSource = ExcelEntities.GetContext().Uslugi.AsEnumerable().OrderBy(x => Convert.ToInt32(x.Id)).ToList();
        }

        private async void ImportBtn_Click(object sender, RoutedEventArgs e)
        {
            string path = System.IO.Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "..", "..", "..", "Template_4335", "Windows", "MuhametzanovaAR", "2.json");
            using (var db = new ExcelEntities())
            {
                var uslugis = await JsonSerializer.DeserializeAsync<List<Uslugi>>(new FileStream(path, FileMode.Open));
                foreach (Uslugi item in uslugis)
                {
                    var uslugi = new Uslugi
                    {
                        Id = item.Id,
                        IdZakaza = item.IdZakaza,
                        DataSozdaniya = item.DataSozdaniya,
                        VremyaZakaza = item.VremyaZakaza,
                        IdClienta = item.IdClienta,
                        Uslugii = item.Uslugii,
                        Statuss = item.Statuss,
                        DataZakritiya = item.DataZakritiya,
                        VremyaProkata = Convert.ToInt32(item.VremyaProkata)
                    };

                    db.Uslugi.Add(uslugi);
                }
                try
                {
                    db.SaveChanges();
                    MessageBox.Show("Данные импортированы!");
                }
                catch (DbEntityValidationException ex)
                {
                    MessageBox.Show(ex.Message);
                }
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

                var app = new Word.Application();
                var document = app.Documents.Add();

                #region Заполнение первой таблицы
                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Категория времени проката 2 часа";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var timeCategories = document.Tables.Add(tableRange, first.Count() + 1, 5);

                    timeCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    timeCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    timeCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = timeCategories.Cell(1, 1).Range;
                    cellRange.Text = "Код";
                    cellRange = timeCategories.Cell(1, 2).Range;
                    cellRange.Text = "Код заказа";
                    cellRange = timeCategories.Cell(1, 3).Range;
                    cellRange.Text = "Дата создания";
                    cellRange = timeCategories.Cell(1, 4).Range;
                    cellRange.Text = "Код элемента";
                    cellRange = timeCategories.Cell(1, 5).Range;
                    cellRange.Text = "Услуги";
                    timeCategories.Rows[1].Range.Bold = 1;
                    timeCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in first)
                    {
                        cellRange = timeCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.Id;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.IdZakaza;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.DataSozdaniya;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 4).Range;
                        cellRange.Text = item.IdClienta;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 5).Range;
                        cellRange.Text = item.Uslugii;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }
                #endregion

                #region Заполнение второй таблицы
                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Категория времени проката 4 часа";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var timeCategories = document.Tables.Add(tableRange, second.Count() + 1, 5);

                    timeCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    timeCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    timeCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = timeCategories.Cell(1, 1).Range;
                    cellRange.Text = "Код";
                    cellRange = timeCategories.Cell(1, 2).Range;
                    cellRange.Text = "Код заказа";
                    cellRange = timeCategories.Cell(1, 3).Range;
                    cellRange.Text = "Дата создания";
                    cellRange = timeCategories.Cell(1, 4).Range;
                    cellRange.Text = "Код элемента";
                    cellRange = timeCategories.Cell(1, 5).Range;
                    cellRange.Text = "Услуги";
                    timeCategories.Rows[1].Range.Bold = 1;
                    timeCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in second)
                    {
                        cellRange = timeCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.Id;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.IdZakaza;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.DataSozdaniya;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 4).Range;
                        cellRange.Text = item.IdClienta;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 5).Range;
                        cellRange.Text = item.Uslugii;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }
                #endregion

                #region Заполнение третьей таблицы
                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Категория времени проката 5,3 часа";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var timeCategories = document.Tables.Add(tableRange, third.Count() + 1, 5);

                    timeCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    timeCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    timeCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = timeCategories.Cell(1, 1).Range;
                    cellRange.Text = "Код";
                    cellRange = timeCategories.Cell(1, 2).Range;
                    cellRange.Text = "Код заказа";
                    cellRange = timeCategories.Cell(1, 3).Range;
                    cellRange.Text = "Дата создания";
                    cellRange = timeCategories.Cell(1, 4).Range;
                    cellRange.Text = "Код элемента";
                    cellRange = timeCategories.Cell(1, 5).Range;
                    cellRange.Text = "Услуги";
                    timeCategories.Rows[1].Range.Bold = 1;
                    timeCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in third)
                    {
                        cellRange = timeCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.Id;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.IdZakaza;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.DataSozdaniya;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 4).Range;
                        cellRange.Text = item.IdClienta;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 5).Range;
                        cellRange.Text = item.Uslugii;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }
                #endregion

                #region Заполнение четвёртой таблицы
                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Категория времени проката 6 часов";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var timeCategories = document.Tables.Add(tableRange, fourth.Count() + 1, 5);

                    timeCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    timeCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    timeCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = timeCategories.Cell(1, 1).Range;
                    cellRange.Text = "Код";
                    cellRange = timeCategories.Cell(1, 2).Range;
                    cellRange.Text = "Код заказа";
                    cellRange = timeCategories.Cell(1, 3).Range;
                    cellRange.Text = "Дата создания";
                    cellRange = timeCategories.Cell(1, 4).Range;
                    cellRange.Text = "Код элемента";
                    cellRange = timeCategories.Cell(1, 5).Range;
                    cellRange.Text = "Услуги";
                    timeCategories.Rows[1].Range.Bold = 1;
                    timeCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in fourth)
                    {
                        cellRange = timeCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.Id;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.IdZakaza;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.DataSozdaniya;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 4).Range;
                        cellRange.Text = item.IdClienta;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 5).Range;
                        cellRange.Text = item.Uslugii;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }
                #endregion

                #region Заполнение пятой таблицы
                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Категория времени проката 8 часов";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var timeCategories = document.Tables.Add(tableRange, fifth.Count() + 1, 5);

                    timeCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    timeCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    timeCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = timeCategories.Cell(1, 1).Range;
                    cellRange.Text = "Код";
                    cellRange = timeCategories.Cell(1, 2).Range;
                    cellRange.Text = "Код заказа";
                    cellRange = timeCategories.Cell(1, 3).Range;
                    cellRange.Text = "Дата создания";
                    cellRange = timeCategories.Cell(1, 4).Range;
                    cellRange.Text = "Код элемента";
                    cellRange = timeCategories.Cell(1, 5).Range;
                    cellRange.Text = "Услуги";
                    timeCategories.Rows[1].Range.Bold = 1;
                    timeCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in fifth)
                    {
                        cellRange = timeCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.Id;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.IdZakaza;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.DataSozdaniya;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 4).Range;
                        cellRange.Text = item.IdClienta;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 5).Range;
                        cellRange.Text = item.Uslugii;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }
                #endregion

                #region Заполнение шестой таблицы
                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Категория времени проката 10 часов";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var timeCategories = document.Tables.Add(tableRange, sixth.Count() + 1, 5);

                    timeCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    timeCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    timeCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = timeCategories.Cell(1, 1).Range;
                    cellRange.Text = "Код";
                    cellRange = timeCategories.Cell(1, 2).Range;
                    cellRange.Text = "Код заказа";
                    cellRange = timeCategories.Cell(1, 3).Range;
                    cellRange.Text = "Дата создания";
                    cellRange = timeCategories.Cell(1, 4).Range;
                    cellRange.Text = "Код элемента";
                    cellRange = timeCategories.Cell(1, 5).Range;
                    cellRange.Text = "Услуги";
                    timeCategories.Rows[1].Range.Bold = 1;
                    timeCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in sixth)
                    {
                        cellRange = timeCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.Id;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.IdZakaza;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.DataSozdaniya;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 4).Range;
                        cellRange.Text = item.IdClienta;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 5).Range;
                        cellRange.Text = item.Uslugii;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }
                #endregion

                #region Заполнение седьмой таблицы
                for (var i = 0; i < 1; i++)
                {
                    var paragraph = document.Paragraphs.Add();
                    var range = paragraph.Range;
                    range.Text = "Категория времени проката 12 часов";
                    paragraph.set_Style("Заголовок 1");
                    range.InsertParagraphAfter();

                    var tableParagraph = document.Paragraphs.Add();
                    var tableRange = tableParagraph.Range;
                    var timeCategories = document.Tables.Add(tableRange, seventh.Count() + 1, 5);

                    timeCategories.Borders.InsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
                    timeCategories.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleDot;
                    timeCategories.Range.Cells.VerticalAlignment = (Word.WdCellVerticalAlignment)Word.WdVerticalAlignment.wdAlignVerticalCenter;

                    Word.Range cellRange;
                    cellRange = timeCategories.Cell(1, 1).Range;
                    cellRange.Text = "Код";
                    cellRange = timeCategories.Cell(1, 2).Range;
                    cellRange.Text = "Код заказа";
                    cellRange = timeCategories.Cell(1, 3).Range;
                    cellRange.Text = "Дата создания";
                    cellRange = timeCategories.Cell(1, 4).Range;
                    cellRange.Text = "Код элемента";
                    cellRange = timeCategories.Cell(1, 5).Range;
                    cellRange.Text = "Услуги";
                    timeCategories.Rows[1].Range.Bold = 1;
                    timeCategories.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                    var count = 1;
                    foreach (var item in seventh)
                    {
                        cellRange = timeCategories.Cell(count + 1, 1).Range;
                        cellRange.Text = item.Id;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 2).Range;
                        cellRange.Text = item.IdZakaza;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 3).Range;
                        cellRange.Text = item.DataSozdaniya;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 4).Range;
                        cellRange.Text = item.IdClienta;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        cellRange = timeCategories.Cell(count + 1, 5).Range;
                        cellRange.Text = item.Uslugii;
                        cellRange.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        count++;
                    }
                }
                #endregion

                app.Visible = true;
            }
        }
    }
}
