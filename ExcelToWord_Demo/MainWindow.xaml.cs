using NPOI.Helper.Excel;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
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
using Newtonsoft.Json;

namespace ExcelToWord_Demo
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public MainWindow()
        {
            InitializeComponent();
        }

        private void Button_Click(object sender, RoutedEventArgs e)
        {
            DataTable dt = ExcelHelper.ExcelConversionDataTable($"{Textbox2.Text}", "");

            var json = JsonConvert.SerializeObject(dt);

            json = json.Replace("\"", "");
            json = json.Replace(",", "\n");
            json = json.Replace("\\n", "\n");
            json = json.Replace("[", "");
            json = json.Replace("]", "");
            json = json.Replace("{", "");
            json = json.Replace("}", "");
            //json = json.Replace("选项F:null","");
            var split = json.Split('\n');
            List<string> list = new List<string>();
            foreach (var item in split)
            {
                if (!item.Contains("null"))
                {
                    list.Add(item);
                }
            }

            string text = "";
            foreach (var item in list)
            {
                text += item;
                text += "\n";
            }

            Textbox1.Text = text;

            //for (int i = 0; i < dt.Rows.Count; i++)
            //{
            //    string strName = dt.Rows[i]["题干"].ToString();
            //}
        }
    }

}
