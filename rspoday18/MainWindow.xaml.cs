using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
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
namespace rspoday18
{
    /// <summary>
    /// Логика взаимодействия для MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public List<Hotel> CurrentHotels { get; set; }
        public MainWindow()
        {
            InitializeComponent();
            var client = new WebClient();
            var response = client.DownloadString("http://localhost:57734/api/hotels");
            CurrentHotels = JsonConvert.DeserializeObject<List<Hotel>>(response);
            DataContext = this;

        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            var application = new Excel.Application();
            Excel.Workbook wb = application.Workbooks.Add(Type.Missing);
            Excel.Worksheet worksheet = wb.Worksheets.Add(Type.Missing);
            worksheet.Cells[1][1] = "Id";
            worksheet.Cells[2][1] = "Name";
            worksheet.Cells[3][1] = "CountOfStars";
            worksheet.Cells[1][1].Font.Bold = worksheet.Cells[2][1].Font.Bold = worksheet.Cells[3][1].Font.Bold = true;
            var client = new WebClient();
            var response = client.DownloadString("http://localhost:59913/api/Hotels");
            CurrentHotels = JsonConvert.DeserializeObject<List<Hotel>>(response);
            int startRowIndex = 2;
            foreach (var hotel in CurrentHotels)
            {
                worksheet.Cells[1][startRowIndex] = hotel.Id;
                worksheet.Cells[2][startRowIndex] = hotel.Name;
                worksheet.Cells[3][startRowIndex] = hotel.CountOfStars;
                startRowIndex++;
            }
            worksheet.Columns.AutoFit();
            application.Visible = true;
            Excel.Range usedRange = worksheet.UsedRange;
            usedRange.HorizontalAlignment = Excel.XlHAlign.xlHAlignCenter;
            usedRange.VerticalAlignment = Excel.XlVAlign.xlVAlignCenter;
        }
    }
}
