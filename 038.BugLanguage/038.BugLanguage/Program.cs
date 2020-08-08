using System;
using System.Threading;

namespace _038.BugLanguage
{
    class Program
    {
        static void Main(string[] args)
        {
        ThreadHead:
            int y = Console.CursorTop;
            string ret = BugLanguage.Convert(Console.ReadLine());
            Console.CursorTop = y;
            Console.WriteLine("                                                       ");
            Console.CursorTop = y;
            Console.WriteLine(ret);
            Thread.Sleep(100);
            goto ThreadHead;
        }
    }
}
