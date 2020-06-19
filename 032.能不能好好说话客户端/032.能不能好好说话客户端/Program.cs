using MSXML2;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace _032.能不能好好说话客户端
{
    class Program
    {
        //致敬：https://lab.magiconch.com/nbnhhsh/

        public static string SeekSX(string word)
        {
            XMLHTTP60 x = new XMLHTTP60();
            x.open("POST", "https://lab.magiconch.com/api/nbnhhsh/guess");
            x.setRequestHeader("content-type", "application/json");
            x.send("{\"text\":\"" + word + "\"}");
            return x.responseText;
            //return x.responseText.Split('[')[2].Split(']')[0];
        }
        static void Main(string[] args)
        {
            Console.WriteLine(SeekSX(Console.ReadLine()));
            Console.ReadLine();
        }
    }
}
