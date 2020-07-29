using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace _037.樱花动漫下载器
{
    class Program
    {
        public static bool loaded = false,loaded2 = false,wrt = false;
        public static int CX, CY;
        [STAThread]
        static void Main(string[] args)
        {
            WebBrowser wb = new WebBrowser();
            wb.DocumentCompleted += Wb_DocumentCompleted;
            Console.Write("Input the url:");
            string url = Console.ReadLine();
            wb.Navigate(url);
            Console.WriteLine("Connecting...");
            loaded = false; do { Application.DoEvents(); } while (!loaded);
            string title = wb.DocumentTitle.Split('_')[0];
            Console.WriteLine("Fetched " + title);
            if (!Directory.Exists(Application.StartupPath + "\\" + title)) Directory.CreateDirectory(Application.StartupPath + "\\" + title);
            List<string> links = new List<string>();
            int start = 1,end = 0;
            foreach (HtmlElement he in wb.Document.GetElementsByTagName("a"))
            {
                string link = null;
                link = he.GetAttribute("href");
                if (link != null)
                {
                    if (link.StartsWith("http://www.imomoe.in/player/") && link.IndexOf("-0-") >= 0 && he.GetAttribute("title").StartsWith("第"))
                    {
                        end++;
                        links.Add(he.GetAttribute("title") + ";" + link);
                        Console.WriteLine(end + ": found " + he.GetAttribute("title") + "!");
                    }
                }
            }
            Console.Write("Download from(1~" + links.Count + "):");
            start = int.Parse(Console.ReadLine());
            Console.Write("Download from(" + start + "~" + links.Count + "):");
            end = int.Parse(Console.ReadLine());
            Console.WriteLine(links.Count + " chapters avaliable.");
            for(int i = start - 1;i < end;i++)
            {
                string link = links[i];
                DateTime d = DateTime.Now;
                Console.WriteLine("Resolving " + link.Split(';')[0] + "...");
                wb.Navigate(link.Split(';')[1]);
                loaded = false; do { Application.DoEvents(); } while (!loaded);
                string ourl = wb.Document.GetElementById("play2").GetAttribute("src");
                ourl = ourl.Split(new string[] { "&vid=" }, StringSplitOptions.None)[1].Split('&')[0];
                Console.WriteLine("Video url got:" + ourl);
                WebClient wc = new WebClient();
                wc.DownloadProgressChanged += Wc_DownloadProgressChanged;
                wc.DownloadFileCompleted += Wc_DownloadFileCompleted;
                Console.WriteLine("Downloading...");
                CX = Console.CursorLeft;CY = Console.CursorTop;
                wc.DownloadFileAsync(new Uri(ourl), Application.StartupPath + "\\" + title + "\\" + link.Split(';')[0] + ".mp4");
                loaded2 = false; do { Application.DoEvents(); } while (!loaded2);
                Console.WriteLine("Completed downloading " + link.Split(';')[0] + ", " + (DateTime.Now - d).TotalSeconds + "s used.");
                wc.Dispose();
            }
            Console.WriteLine("\nFinished all tasks b（￣▽￣）d　!");
            Console.ReadLine();
        }

        private static void Wc_DownloadFileCompleted(object sender, System.ComponentModel.AsyncCompletedEventArgs e)
        {
            loaded2 = true;
        }

        private static void Wc_DownloadProgressChanged(object sender, DownloadProgressChangedEventArgs e)
        {
            if (loaded2) return;
            if (wrt) return;
            wrt = true;
            Console.SetCursorPosition(CX, CY);
            Console.WriteLine(e.BytesReceived + " bytes / " + e.TotalBytesToReceive + " bytes (" + (e.ProgressPercentage + 1) + "%) ...");
            Console.SetCursorPosition(CX, CY);
            wrt = false;
        }

        private static void Wb_DocumentCompleted(object sender, WebBrowserDocumentCompletedEventArgs e)
        {
            loaded = true;
        }
    }
}
