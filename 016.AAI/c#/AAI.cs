// Artifical Artifical Intelligence
// 虚假AI
// 作者：Buger404

using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Windows.Forms;
using System.Web;
using System.Threading.Tasks;
using System.Collections;
using System.Threading;
using System.IO;
using MSXML2;

namespace ArtificalA.Intelligence
{
    public class ArtificalAI
    {
        public static bool DebugLog = false;
        public static int recordtime = 0;
        private static void Log(string log, ConsoleColor color = ConsoleColor.White)
        {
            if (DebugLog == false) { return; }
            if (recordtime < DateTime.Now.Hour) { recordtime = DateTime.Now.Hour; log = "##TIMELINE(" + recordtime + ")####################################################\r\n" + log; }
            try
            {
                File.AppendAllText(@"C:\DataArrange\Log\[AAI]-" + MainThread.MessagePoster.logid + ".txt", log + "\r\n");
            }
            catch
            {

            }
            Console.ForegroundColor = color;
            Console.WriteLine(log);
        }

        public static string Talk(string Question, string engine)
        {
            string ret = ""; string url = ""; string q = WebUtility.UrlEncode(Question);
            switch (engine)
            {
                case ("baidu"):
                    url = "https://zhidao.baidu.com/search?lm=0&rn=10&pn=0&fr=search&ie=gbk&word=" + q;
                    ret = Search(url, @"<a[^>]*?href=[^>]*?class=""ti""[^>]*?>", @"<div[^>]*?id=""[^>]*?-content-[^>]*?>[^>]*?<div[^>]*?>[^>]*?<div[^>]*?>[^>]*?[^>]*?<span[^>]*?>[^>]*?</span>[^>]*?</div>[^>]*?</div>[^>]*?.*?[^>]*?</div>", engine);
                    break;
                case ("tieba"):
                    url = "http://tieba.baidu.com/f/search/res?ie=utf-8&qw=" + q + "&red_tag=e0313206225";
                    ret = Search(url, @"<a[^>]*?class=""bluelink""[^>]*?target=""_blank""[^>]*?>", @"div[^>]*id=""post_[^>]*>.*?</div>", engine);
                    break;
                case ("csdn"):
                    url = "https://so.csdn.net/so/search/s.do?q=" + q + "&t=&u=";
                    ret = Search(url, @"<a[^>]*?href=""[^>]*?target=""_blank""[^>]*?>", @"div[^>]*id=""post_[^>]*>.*?</div>", engine);
                    break;

                case ("msdn"):
                    //ret = Search(Question, engine,
                    //   "https://docs.microsoft.com/zh-cn/search/?search={q}&category=All&scope=Desktop",
                    //    "a", "searchItem.0", "data-bi-name", "href",
                    //    "main", "main", "id", "main", "MSDN无相关结果");
                    break;
            }
            return ret;
        }

        private static string GetHTML(string url)
        {
            Log("Connect:" + url, ConsoleColor.Yellow);
            XMLHTTP x = new XMLHTTP();
            x.open("GET", url, false);
            x.send();
            Log("Web data received !", ConsoleColor.Yellow);
            Byte[] b = (Byte[])x.responseBody;
            string s = System.Text.ASCIIEncoding.UTF8.GetString(b, 0, b.Length);
            return s;
        }

        private static string GetInner(string code, int buff = 0)
        {
            string[] t = code.Replace('<', '>').Split('>'); string r = "";

            //过滤除br外的所有标签
            for (int i = 1; i < t.Length; i++)
            {
                if (i % 2 == buff)
                {
                    if (t[i] == "\br") { r += "\n"; }
                }
                else
                {
                    r += t[i];
                }
            }
            return r;
        }

        private static string GetAttr(string code, string name)
        {
            return code.ToLower().Split(new[] { name + "=\"" }, StringSplitOptions.None)[1].Split('\"')[0];
        }

        private static string Search(string url, string linkexp, string conexp, string engine)
        {
            string code = GetHTML(url); Random r = new Random(); bool choice = false;
            List<string> links = new List<string>(); string l = "";

            foreach (Match m in Regex.Matches(code, linkexp))
            {
                l = GetAttr(m.Value, "href"); choice = true;
                if (engine == "baidu") { l = l.Replace("http:", "https:"); }
                if (engine == "tieba") { choice = (l.IndexOf("http") != 0); }
                if (choice) { Log("Link:" + l); links.Add(l); }
            }
            if (links.Count == 0) { Log("Faile to get links !", ConsoleColor.Red); return ""; }
            string link = links[r.Next(0, links.Count)];
            if (engine == "tieba") { link = "https://tieba.baidu.com/" + link; }
            List<string> rs = new List<string>();
            code = GetHTML(link);
            if (engine == "baidu") { code = code.Replace("\n", ""); }

            foreach (Match m in Regex.Matches(code, conexp))
            {
                if (engine == "baidu") { l = GetInner(m.Value, 1).Trim(); }
                if (engine == "tieba") { l = GetInner(m.Value).Trim(); }
                //Log("Content:" + l.Length + " words");
                if (l.Length > 0) { rs.Add(l); }
            }
            if (rs.Count == 0) { Log("Faile to get contents !", ConsoleColor.Red); return ""; }
            string tts = rs[r.Next(0, rs.Count)];
            if(tts.ToLower().IndexOf("http:") >= 0 || tts.ToLower().IndexOf("https:") >= 0 || tts.ToLower().IndexOf("#") >= 0 || tts.ToLower().IndexOf("qq") >= 0)
            {
                Log("AD:" + tts, ConsoleColor.Red); tts = "";
            }
                
            Log("Reply:" + tts, ConsoleColor.Green);

            return tts;
        }

    }
}
