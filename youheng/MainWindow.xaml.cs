using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using ExcelEasyUtil;
using Path = System.IO.Path;
using System.Threading;
using System.Windows.Threading;

namespace youheng
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {

        public List<string> ListTxt { get; set; }

        public List<Rowa> ListRowa { get; set; }
        public MainWindow()
        {
            InitializeComponent();
        }

        private void btnAdd_Click(object sender, RoutedEventArgs e)
        {
            readFIle();
            if (ListTxt.Count > 0)
            {
                ToDataClass();
            }
            txtlabel.Content = $"共计找到【{ListTxt.Count}】条记录\r\n其中有效数据【{ListRowa.Count}】条";
        }

        public void readFIle()
        {
            ListTxt = new List<string>();
            OpenFileDialog openFile = new OpenFileDialog();
            openFile.Filter = "文本文件(*.txt)|*.txt";
            if (openFile.ShowDialog() == true)
            {
                using (StreamReader sr = new StreamReader(openFile.FileName, Encoding.Default))
                {
                    int lineCount = 0;
                    while (sr.Peek() > 0)
                    {
                        lineCount++;
                        string temp = sr.ReadLine();
                        if (temp != "")
                            ListTxt.Add(temp);
                    }
                }
                this.textBox.Text = "准备显示分析转换结果";
                this.pro1.Value = 0;
            }
        }

        private void btnExl_Click(object sender, RoutedEventArgs e)
        {
            if (ListRowa == null || ListRowa.Count == 0)
            {
                this.textBox.Text = "没有导入数据源txt文档或没找到有效数据\r\n 请重新选择数据源";
                
            }
            else
            {
                //showList();
                update();
                this.btnExl.IsEnabled = false;
                saveList();
            }

        }


        private void update()
        {
            this.textBox.Text = "";
            var h = new Thread(() =>
            {
           var v = Math.Round(((double)100 / ListRowa.Count), 2);

                ListRowa.ForEach(r => {
                    this.Dispatcher.BeginInvoke((Action)(() =>
                    {
                        this.textBox.AppendText($"【{r.Name}】【{r.Id}】{r.Lianxi }\r\n");
                        this.pro1.Value += v;
                    }));
                    Thread.Sleep(200);
                });
                this.Dispatcher.BeginInvoke((Action)(() =>
                {
                    this.btnExl.IsEnabled = true;
                }));
            });
            h.Start();
            
        }

        private void ToDataClass()
        {
            ListRowa = new List<Rowa>();
            foreach (string item in ListTxt)
            {
                string[] strArr = item.Split(new[] { "----" }, StringSplitOptions.None);
                if (strArr.Length == 3)
                {
                    if (strArr[2] != "" && strArr[2] != "×")
                    {
                        var v = new Rowa();
                        v.Name = strArr[0];
                        v.Id = strArr[1];
                        v.TypeName = null;
                        v.Lianxi = $"信息：{strArr[2]}";
                        ListRowa.Add(v);
                    }
                }
                else if (strArr.Length == 5)
                {
                    if (strArr[2] != "×" || strArr[3] != "×"|| strArr[4] != "×")
                    {
                        var v = new Rowa();
                        v.Name = strArr[0];
                        v.Id = strArr[1];
                        v.TypeName = null;
                        string st = "";
                        strArr.Skip(2).Where(r => r.Length > 1).ToList()
                            .ForEach(r => st+=r);
                        v.Lianxi = $"信息：【{st}】";
                        ListRowa.Add(v);
                    }
                }
            }
        }

        public void saveList() 
        {
            string savedPath = Core.CreateXlsWorkBook()
                .FillSheet("查询信息导入模板", ListRowa,//填充第一个表格
                                                //new Dictionary<string, Expression<Func<People, dynamic>>> //设置表格头，原始类型
               new PropertyColumnMapping<Rowa> //设置表格头，专用简化类型
               {
                {"证件号",p=>p.Id },{"案人姓名",p=>p.Name },{"信息类型",p=>p.TypeName},{"信息内容",p=>p.Lianxi },{"备注",p=>p.Beizhu }
               })
                .SaveToFile(Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "案人数据导入模板.xls"));
        }

            
        
    }

}
