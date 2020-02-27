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
using Newtonsoft.Json.Linq;
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
            saveAs.Filter = "Excel files(*.xlsx)|*.xlsx";
            saveAs.AddExtension = true;
            saveAs.DefaultExt = ".xlsx";
            saveAs.OverwritePrompt = true;
            saveAs.Title = "Where do you want to save your spreadsheet?";

            if (saveAs.ShowDialog() == true)
            {
                string location = saveAs.FileName;

                Excel.Application xlApp = new Excel.Application();
                Excel.Workbook wbResult;
                Excel.Worksheet resultSheet;
                Excel.Range wbRange;
                
                //xlApp.Visible = true;

                wbResult = xlApp.Workbooks.Add();
               // wbResult.Application.Visible = true;

                resultSheet = wbResult.ActiveSheet;
                wbRange = resultSheet.Range["A1","H1"];
                string[] headers = {"Name", "TwitterHandle", "Description", "Location","URL","Email","FollowerNo","FollowingNo" };
                int i = 0;
                foreach (Excel.Range item in wbRange)
                {
                    item.Value = headers[i];
                    i++;
                }
                wbRange.Interior.Color = Excel.XlRgbColor.rgbLightGrey;
                wbRange.Font.Bold = true;
                wbRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].LineStyle = Excel.XlLineStyle.xlContinuous;
                wbRange.Borders[Excel.XlBordersIndex.xlEdgeBottom].Weight = 1d;

                //Twitter api calls
                Auth.SetUserCredentials("7HvRXvXPf9pt4EaEXaKUf24Cc",
                        "IYgU2IZa74bBEatM4Q41dNQN0RXVWetyFrJZlPWBxEUfx8RR7e",
                        "781216822116937728-bW9jR4ZnMkZBV1qIkS8GWgkLkQIrX8L",
                        "DK1QD9CsIJuvgkccw4lefVrA7Jb4SmWjvbOzZ7QDh68uR");

                var handle = User.GetUserFromScreenName("JamesASinclair");
                var followers = User.GetFollowerIds(handle, 10000);
                var followerDetails = User.GetUsersFromIds(followers);

                string fullJson = followerDetails.ToJson();
                //wbRange.Value = followers.ToJson();

                JToken specificJson = JToken.Parse(fullJson);

                //get the parsed result into a list
                IList<JToken> results = specificJson.Children().ToList();

                //serialize into objects in sResults list
                IList<SearchResult> sResults = new List<SearchResult>();
                foreach (JToken result in results)
                {
                    SearchResult searchResult = result.ToObject<SearchResult>();
                    sResults.Add(searchResult);
                }

                Excel.Range resultRange = resultSheet.Range["A2"];
                for (int r = 0; r < sResults.Count; r++)
                {
                    for (int c = 0; c < 7; c++)
                    {
                        string toAdd = "";
                        switch (c)
                        {
                            case 0: //Name
                                toAdd = sResults[r].name;
                                break;
                            case 1: //TwitterHandle
                                toAdd = sResults[r].screen_name;
                                break;
                            case 2: //Description
                                toAdd = sResults[r].description;
                                break;
                            case 3: //Location
                                toAdd = sResults[r].location;
                                break;
                            case 4: //url
                                toAdd = sResults[r].url;
                                break;
                            case 5: //followno
                                toAdd = sResults[r].followers_count;
                                break;
                            case 6: //followingno
                                toAdd = sResults[r].friends_count;
                                break;
                            default:
                                break;
                        }
                            
                        resultRange.Offset[r, c].Value = toAdd;
                    }
                    
                }

                resultRange.Columns["A:H"].AutoFit();
                xlApp.DisplayAlerts = false;
                wbResult.SaveAs(Filename: location);
                xlApp.DisplayAlerts = true;
                wbResult.Close();
                xlApp.Quit();

            } 

            
        }
        public class SearchResult
        {
            public string name { get; set; }
            public string description { get; set; }
            public string location { get; set; }
            public string url { get; set; }
            public string followers_count { get; set; }
            public string friends_count { get; set; }
            public string screen_name { get; set; }
        }
    }
}
