using System;
using System.Text.RegularExpressions;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Web;
using System.Threading.Tasks;
using System.Collections;
using System.Threading;
using System.IO;
using Winista.Text.HtmlParser;
using MSXML2;
using Winista.Text.HtmlParser.Util;
using Winista.Text.HtmlParser.Filters;
using Winista.Text.HtmlParser.Lex;

namespace ConsoleWeb
{
    class Program
    {
        private static bool LINK = false;
        private static void Log(string log, ConsoleColor color = ConsoleColor.White)
        {
            Console.ForegroundColor = color;
            Console.WriteLine(log);
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

        private static void Outputs(NodeList nodes)
        {
            for (int i = 0; i < nodes.Count; i++)
            {
                INode node = nodes[i];
                if (node is ITag)
                {
                    ITag tag = (ITag)node;
                    if (tag.TagName == "DIV" || tag.TagName == "BR" || tag.TagName == "TR" || tag.TagName == "H3" || tag.TagName == "P")
                    {
                        Console.WriteLine("");
                    }
                    if (tag.Children != null)
                    {
                        if (tag.TagName != "STYLE" && tag.TagName != "SCRIPT" && tag.TagName != "NOSCRIPT")
                        {
                            //Console.Write(tag.TagName + " ");
                            if (tag.TagName == "LINK" || tag.Attributes["HREF"] != null) 
                                Console.ForegroundColor = ConsoleColor.Green;
                            Outputs(tag.Children);
                            Console.ForegroundColor = ConsoleColor.Gray;
                        }
                    }

                    /**.WriteLine(tag.TagName);
                    if (tag.Attributes["ID"] != null)
                    {
                        Console.WriteLine("{ id=\"" + tag.Attributes["ID"].ToString() + "\" }");
                    }
                    if (tag.Attributes["HREF"] != null)
                    {
                        Console.WriteLine("{ href=\"" + tag.Attributes["HREF"].ToString() + "\" }");
                    }**/

                }
                else
                {
                    if(node is IText)
                    {
                        if (node.Children != null)
                        {
                            Outputs(node.Children);
                        }
                        else
                        {
                            Console.Write(node.GetText().Trim().Replace("&nbsp;"," ") + " ");
                        }
                    }
                    
                }
            }
        }

        static void Main(string[] args)
        {
            string url = "http://www.baidu.com/s?ie=UTF-8&wd=WDDM";
            //string code = GetHTML(url);
            System.Net.WebClient web = new System.Net.WebClient();
            web.Encoding = System.Text.Encoding.UTF8;
            string code = web.DownloadString(url);

            Lexer l = new Lexer(code); Parser p = new Parser(l);
            NodeList rnode = p.Parse(null);
            Outputs(rnode);

            string cmd = Console.ReadLine();
        }
    }
}
