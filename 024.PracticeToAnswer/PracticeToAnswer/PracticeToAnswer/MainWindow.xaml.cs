using MSXML2;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Forms;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;

namespace PracticeToAnswer
{
    /// <summary>
    /// MainWindow.xaml 的交互逻辑
    /// </summary>
    public partial class MainWindow : Window
    {
        public int Loaded = 0;
        public List<String> Ques = new List<string>();
        public List<String> Ans = new List<string>();
        public int qin = 0;
        public System.Windows.Forms.WebBrowser w = new System.Windows.Forms.WebBrowser();
        public MainWindow()
        {
            InitializeComponent();
            w.DocumentCompleted += W_DocumentCompleted;
            StateText.Content = "All ready";
            w.ScriptErrorsSuppressed = false;
        }
        private static string GetHTML(string url)
        {
            XMLHTTP x = new XMLHTTP();
            x.open("GET", url, false);
            x.send();
            Byte[] b = (Byte[])x.responseBody;
            string s = System.Text.ASCIIEncoding.UTF8.GetString(b, 0, b.Length);
            return s;
        }

        public void GetAnswer()
        {
            Ques.Clear();Ans.Clear();
            string buff = "";string num = "";bool added = true;
            for(int i = 0;i < Inputs.Text.Length; i++) 
            {
                added = true;
                if(Inputs.Text[i] == '．' || Inputs.Text[i] == '.')
                {
                    string[] t = buff.Split('\n');
                    try
                    {
                        if (Convert.ToInt32(t[t.Length - 1]).ToString() == t[t.Length - 1])
                        {
                            added = false;
                            if (num != "")
                            {
                                Ques.Add(buff);
                                Outputs.Text = Outputs.Text +
                                               "/////////////////////////////////////\n" +
                                               "No. " + num + "\n" +
                                               "/////////////////////////////////////\n" +
                                               buff + "\n" +
                                               "/////////////////////////////////////\n";
                            }
                            buff = ""; num = t[t.Length - 1];
                        }
                    }
                    catch
                    {
                        added = true;
                    }

                }

                if (added) { buff += Inputs.Text[i]; }

            }

            if (num != "" && buff != "")
            {
                Ques.Add(buff);
                Outputs.Text = Outputs.Text +
                               "/////////////////////////////////////\n" +
                               "No. " + num + "\n" +
                               "/////////////////////////////////////\n" +
                               buff + "\n" +
                               "/////////////////////////////////////\n";
            }

            qin = 0;
            FetchAnswer();
        }
        public void FetchAnswer()
        {
            for (int i = 1; i <= 10; i++)
            {
                System.Windows.Forms.Application.DoEvents();
                Thread.Sleep(50);
            }
            if (qin >= Ques.Count)
            {
                string r = "";
                for(int i = 0;i < Ans.Count; i++)
                {
                    r = r + "-------------------------------------------\n" +
                            "第" + (i + 1) + "题 解析\n" +
                            Ans[i] +  "\n" +
                            "-------------------------------------------\n";
                }
                Outputs.Text = r;
                StateText.Content = "Answers complete";
                return;
            }
            Loaded = 0;
            StateText.Content = "Waiting ...";
            while (w.IsBusy)
            {
                System.Windows.Forms.Application.DoEvents();
            }
            StateText.Content = "Connecting " + qin + "/" + Ques.Count + " ...";
            w.Navigate("https://www.baidu.com/s?ie=utf-8&f=8&rsv_bp=1&tn=baidu" +
                    "&wd=" + System.Web.HttpUtility.UrlEncode(Ques[qin]) +
                    "&oq=%25E4%25BD%259C%25E4%25B8%259A%25E5%25B8%25AE%25E7%25BD%2591%25E9%25A1%25B5%25E7%2589%2588" +
                    "&rsv_pq=bc5cb5420006236a&rsv_t=ae75FPPDoOSilg2QWOJLpQcV7BA8IHKTLr4DSyZz%2FCcQ2hCiya%2BvGZtJvGk" +
                    "&rqlang=cn&rsv_enter=1&rsv_dl=tb&inputT=22915" +
                    "&si=www.zybang.com&ct=2097152");
            qin++;

        }
        private void W_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            Loaded++;
            switch (Loaded)
            {
                case (1):
                    HtmlDocument doc = (HtmlDocument)w.Document;
                    List<String> links = new List<String>();
                    links.Clear();
                    foreach(HtmlElement l in doc.Links)
                    {
                        StateText.Content = l.GetAttribute("href");
                        if (l.GetAttribute("href").StartsWith("http://www.baidu.com/link"))
                        {
                            Outputs.Text = Outputs.Text + "\nurl:" + l.GetAttribute("href");
                            links.Add(l.GetAttribute("href"));
                        }
                    }
                    if(links.Count == 0)
                    {
                        StateText.Content = "Unknown";
                        Ans.Add("不知道");
                        FetchAnswer();
                        return;
                    }
                    w.Stop();
                    int fail = 0;
                relink:
                    for (int i = 1;i <= 10; i++)
                    {
                        System.Windows.Forms.Application.DoEvents();
                        Thread.Sleep(50);
                    }
                    try
                    {
                        StateText.Content = "Waiting ...";
                        while (w.IsBusy)
                        {
                            System.Windows.Forms.Application.DoEvents();
                        }
                        w.Navigate(links[1]);
                    }
                    catch
                    {
                        fail++;
                        if(fail > 3)
                        {
                            StateText.Content = "Unknown";
                            Ans.Add("不知道");
                            FetchAnswer();
                            return;
                        }
                        goto relink;
                    }
                    StateText.Content = links[1];
                    break;
                default:
                    Outputs.Text = "Jumping:" + Loaded + "\nurl:" + w.Url.ToString() + "\nsource:\n" + w.DocumentText ;
                    break;
                case (2):
                    HtmlElement h = w.Document.GetElementById("good-answer");
                    if (h == null)
                    {
                        StateText.Content = "Unknown";
                        Ans.Add(w.Document.Body.InnerText);
                        FetchAnswer();
                        return;
                    }
                    h = h.Children[1].Children[0];
                    Outputs.Text = h.InnerText;
                    Ans.Add(h.InnerText);
                    //foreach(HtmlElement h in w.Document.GetElementsByTagName("dl"))
                    //{
                    //    if(h.GetAttribute("id") == "good-answer")
                    //    {
                    //        h.FirstChild.FirstChild
                    //        Outputs.Text = h.InnerText;
                    //    }
                    //}
                    StateText.Content = "Complete " + qin + "/" + Ques.Count;
                    FetchAnswer();
                    break;
            }
        }

        private void ConBtn_Click(object sender, RoutedEventArgs e)
        {
            GetAnswer();
        }

        private void ShowBtn_Click(object sender, RoutedEventArgs e)
        {
            string r = "";
            for (int i = 0; i < Ans.Count; i++)
            {
                r = r + "-------------------------------------------\n" +
                        "第" + (i + 1) + "题 解析\n" +
                        Ans[i] + "\n" +
                        "-------------------------------------------\n";
            }
            Outputs.Text = r;
            StateText.Content = "Answers complete";
            return;
        }
    }
}
