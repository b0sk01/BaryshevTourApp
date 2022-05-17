using Microsoft.VisualBasic.ApplicationServices;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms.DataVisualization.Charting;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using Excel = Microsoft.Office.Interop.Excel;
using Word = Microsoft.Office.Interop.Word;

namespace BaryshevPaymentsExampleApp
{  
    public partial class MainWindow : Window
    {
        private BaryshevPaymentsBaseEntities4 _context = new BaryshevPaymentsBaseEntities4(); //Создаем переменную
        public MainWindow()
        {
            InitializeComponent();
            ChartPayments.ChartAreas.Add(new ChartArea("Main")); //Cоздаем новую коллекцию 

            var currentSeries = new Series("Payments") //создаем объект
            {
                IsValueShownAsLabel = true //делаем видимым в диаграмме
            };
            ChartPayments.Series.Add(currentSeries); //Добавляем объект в коллекцию 

            ComboUsers.ItemsSource = _context.Users.ToList(); //загрузим значения из таблицы Users в выпадающий список
            ComboChartTypes.ItemsSource = Enum.GetValues(typeof(SeriesChartType)); //Получим типы диаграмм из перечисления SeriesChartType
        }

        private void UpdateChart(object sender, SelectionChangedEventArgs e)
        {
            if (ComboUsers.SelectedItem is User currentUser && ComboChartTypes.SelectedItem is SeriesChartType currentType) //Получим выведенные значения как currentUser и currentType
            {
                Series currentSeries = ChartPayments.Series.FirstOrDefault(); //Получаем серию данных из коллекции
                currentSeries.ChartType = currentType; //Устанавливем тип диаграммы
                currentSeries.Points.Clear(); //Очищаем предыдущие значения

                var categoriesList = _context.Categories.ToList(); //Получим список категорий из базы данных
                foreach (var category in categoriesList)
                {
                    currentSeries.Points.AddXY(category.Name, _context.Payments.ToList().Where(p => p.User == currentUser && p.Category == category).Sum(p => p.Price * p.Num));
                }    
            }
        }
        private void BtnExportToExcel_Click(object sender, RoutedEventArgs e)
        {
            var allUsers = _context.Users.ToList().OrderBy(p => p.FIO).ToList(); //Загружаем данные из таблицы Users и сортируем их по ФИО

            var application = new Excel.Application(); //Объявляем переменную с приложение Excel
            application.SheetsInNewWorkbook = allUsers.Count(); //Установим кол-во листов равное кол-ву пользователей в БД

            Excel.Workbook workbook = application.Workbooks.Add(Type.Missing); //Добавим книгу Excel Workbook

            int startRowIndex = 1; //устанавливаем первоначальное значение переменной 

            for (int i = 0; i < allUsers.Count(); i++) //заполняем данные каждого листа
            {
                Excel.Worksheet worksheet = application.Worksheets.Item[i + 1]; //Получаем текущий лист по индексу
                worksheet.Name = allUsers[i].FIO; //Назначаем листу имя ФИО пользователя

                worksheet.Cells[1][startRowIndex] = "Дата платежа"; //укажем в верхней строке заголовки столбцов
                worksheet.Cells[2][startRowIndex] = "Название";
                worksheet.Cells[3][startRowIndex] = "Стоимость";
                worksheet.Cells[4][startRowIndex] = "Количество";
                worksheet.Cells[5][startRowIndex] = "Сумма";

                startRowIndex++;

                var usersCategories = allUsers[i].Payments.OrderBy(p => p.Date).GroupBy(p => p.Category).OrderBy(p => p.Key.Name); //Группируем платежи пользователя по категориям, отсортируем по дате,
                                                                                                                                   //категроии и ключевому имени

                foreach (var groupCategory in usersCategories) //проходим по списку категории 
                {
                    Excel.Range headerRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[5][startRowIndex]]; //Получаем диапазон для заголовка
                    headerRange.Merge(); //Заполняем значения
                    headerRange.Value = groupCategory.Key.Name; //Объединяем ячейки
                    headerRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter; //Делаем горизонатльное выравнивание по центру
                    headerRange.Font.Italic = true; //применяем курсив

                    startRowIndex++;

                    foreach (var payment in groupCategory) //Проходим по таблице Payment
                    {
                        worksheet.Cells[1][startRowIndex] = payment.Date.ToString("dd.MM.yyyy HH:mm"); //Заполняем по столбцу и строке данные
                        worksheet.Cells[2][startRowIndex] = payment.Name;
                        worksheet.Cells[3][startRowIndex] = payment.Price;
                        worksheet.Cells[4][startRowIndex] = payment.Num;

                        worksheet.Cells[5][startRowIndex].Formula = $" =C{startRowIndex}*D{startRowIndex}"; //считаем по формуле данные

                        worksheet.Cells[3][startRowIndex].NumberFormat =
                            worksheet.Cells[3][startRowIndex].NumberFormat = "####"; //Устанавливаем цийровой формат

                        startRowIndex++;
                    }

                    Excel.Range sumRange = worksheet.Range[worksheet.Cells[1][startRowIndex], worksheet.Cells[4][startRowIndex]]; //Выбираем диапазон для объединения ячеек ИТОГО
                    sumRange.Merge(); //Заполняем значения
                    sumRange.Value = "ИТОГО:";
                    sumRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignRight; //Выравниваем по правому краю 

                    worksheet.Cells[5][startRowIndex].Formula = $" =SUM(E{startRowIndex - groupCategory.Count()}:)" + $" E{startRowIndex - 1})"; //Считаем суммы внутри категории

                    sumRange.Font.Bold = worksheet.Cells[5][startRowIndex].Font.Bold = true; //Делаем текст жирным
                    worksheet.Cells[5][startRowIndex].NumberFormat = "####"; //применяем цифровой формат

                    startRowIndex++;

                    Excel.Range rangeBorders = worksheet.Range[worksheet.Cells[1][1], worksheet.Cells[5][startRowIndex - 1]]; //Делаем границы для таблицы
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeLeft].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeRight].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlEdgeTop].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideHorizontal].LineStyle =
                    rangeBorders.Borders[Excel.XlBordersIndex.xlInsideVertical].LineStyle = Excel.XlLineStyle.xlContinuous;

                    worksheet.Columns.AutoFit(); //автоматический подбор ширины для всех столбцов 
                }
            }
            application.Visible = true; //Отобразим приложение
        }

        private void BtnExportToWord_Click(object sender, RoutedEventArgs e)
        {
            var allUsers = _context.Users.ToList();
            var allCategories = _context.Categories.ToList();

            var application = new Word.Application();

            Word.Document document = application.Documents.Add();

            foreach (var user in allUsers)
            {
                Word.Paragraph userParagraph = document.Paragraphs.Add();
                Word.Range userRange = userParagraph.Range;
                userRange.Text = user.FIO;
                userParagraph.set_Style("Заголовок");
                userRange.InsertParagraphAfter();

                Word.Paragraph tableParagraph = document.Paragraphs.Add();
                Word.Range tableRange = tableParagraph.Range;
                Word.Table paymentsTable = document.Tables.Add(tableRange, allCategories.Count() + 1, 3);
                paymentsTable.Borders.InsideLineStyle = paymentsTable.Borders.OutsideLineStyle
                = Word.WdLineStyle.wdLineStyleSingle;
                paymentsTable.Range.Cells.VerticalAlignment = Word.WdCellVerticalAlignment.wdCellAlignVerticalCenter;

                Word.Range cellRange;

                cellRange = paymentsTable.Cell(1, 1).Range;
                cellRange.Text = "Иконка";
                cellRange = paymentsTable.Cell(1, 2).Range;
                cellRange.Text = "Категория";
                cellRange = paymentsTable.Cell(1, 3).Range;
                cellRange.Text = "Сумма расходов";

                paymentsTable.Rows[1].Range.Bold = 1;
                paymentsTable.Rows[1].Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;

                for (int i = 0; i < allCategories.Count(); i++)
                {
                    var currentCategory = allCategories[i];

                    cellRange = paymentsTable.Cell(i + 2, 1).Range;
                    string path = @"Z:\BaryshevPaymentsExampleApp\Assests\dragon.png";
                    Console.WriteLine(path);
                    Word.InlineShape imageShape = cellRange.InlineShapes.AddPicture(path);

                    cellRange = paymentsTable.Cell(i + 2, 2).Range;
                    cellRange.Text = currentCategory.Name;

                    cellRange = paymentsTable.Cell(i + 2, 3).Range;
                    cellRange.Text = user.Payments.ToList()
                    .Where(p => p.Category == currentCategory).Sum(p => p.Num * p.Price).ToString("N2") + " руб.";

                }

                Payment maxPayment = user.Payments
                .OrderByDescending(p => p.Price * p.Num).FirstOrDefault();
                if (maxPayment != null)
                {
                    Word.Paragraph maxPaymentParagraph = document.Paragraphs.Add();
                    Word.Range maxPaymentrange = maxPaymentParagraph.Range;
                    maxPaymentrange.Text = $"Самы дорогостоющий платеж - {maxPayment.Name} за {(maxPayment.Price * maxPayment.Num).ToString("N2")}" +
                    $"руб. от {maxPayment.Date.ToString("dd.MM.yyyy HH:mm")}";
                    maxPaymentParagraph.set_Style("Цитата 2");
                    maxPaymentrange.Font.Color = Word.WdColor.wdColorDarkRed;
                    maxPaymentrange.InsertParagraphAfter();
                }

                Payment minPayment = user.Payments
                .OrderBy(p => p.Price * p.Num).FirstOrDefault();
                if (minPayment != null)
                {
                    Word.Paragraph minPaymentParagraph = document.Paragraphs.Add();
                    Word.Range minPaymentrange = minPaymentParagraph.Range;
                    minPaymentrange.Text = $"Самы дешевый платеж - {minPayment.Name} за {(minPayment.Price * minPayment.Num).ToString("N2")}" +
                    $"руб. от {maxPayment.Date.ToString("dd.MM.yyyy HH:mm")}";
                    minPaymentParagraph.set_Style("Цитата 2");
                    minPaymentrange.Font.Color = Word.WdColor.wdColorGreen;

                }

                if (user != allUsers.LastOrDefault())
                    document.Words.Last.InsertBreak(Word.WdBreakType.wdPageBreak);
            }

            application.Visible = true;

            document.SaveAs2(@"C:\WordPayment\Test.docx");
            document.SaveAs2(@"C:\WordPayment\Test.pdf", Word.WdExportFormat.wdExportFormatPDF);
        }
    }
}
