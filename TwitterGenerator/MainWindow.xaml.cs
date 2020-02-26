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
using Microsoft.Win32;
using Tweetinvi;

namespace TwitterGenerator
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

        private void BtnGenerate_Click(object sender, RoutedEventArgs e)
        {
            //SAVE AS FIRST
            SaveFileDialog saveAs = new SaveFileDialog();
            saveAs.Filter = "Excel files(*.xls or.xlsx)|.xls; *.xlsx";
            saveAs.OverwritePrompt = true;
            saveAs.Title = "Where do you want to save your spreadsheet?";

            if (saveAs.ShowDialog() == true)
            {
                string location = saveAs.FileName;

                //Auth.SetUserCredentials("7HvRXvXPf9pt4EaEXaKUf24Cc",
                //                        "IYgU2IZa74bBEatM4Q41dNQN0RXVWetyFrJZlPWBxEUfx8RR7e",
                //                        "781216822116937728-bW9jR4ZnMkZBV1qIkS8GWgkLkQIrX8L",
                //                        "DK1QD9CsIJuvgkccw4lefVrA7Jb4SmWjvbOzZ7QDh68uR");

                //var handle = User.GetUserFromScreenName("dailyskyfox");
                //var followers = User.GetFollowers(handle, 10);

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook wbResult = new Excel.Workbook();
                Excel.Range wbRange;
                xlApp.Visible = true;
                wbResult = xlApp.Workbooks.Add();
                wbResult.Application.Visible = true;
                wbRange = wbResult.Application.Range(0, 0);



            } 


            
        }
    }
}
