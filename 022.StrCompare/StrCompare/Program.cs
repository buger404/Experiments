using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace StrCompare
{
    class Program
    {
        public static double CompareStr(string s1,string s2)
        {
            int c1 = 0;string s;string s3;
            s = (s1.Length > s2.Length ? s1 : s2);
            s3 = (s1.Length < s2.Length ? s1 : s2);
            for (int i = 0;i < s3.Length; i++)
            {
                for(int j = 0;j < s.Length; j++)
                {
                    if(s[j] == s3[i])
                    {
                        s.Remove(j, 1);
                        c1++;break;
                    }
                }
            }
            double ret = (s3.Length * 1f / s.Length * 1f) * 0.3 + (c1 * 1f / s.Length * 1f) * 0.7;
            ret = Math.Pow(ret * 2, 2);
            return Math.Pow(ret / 2,2);
        }
        static void Main(string[] args)
        {
            string s1 = Console.ReadLine();
            string s2 = Console.ReadLine();
            Console.WriteLine(CompareStr(s1, s2));
            Main(args);
        }
    }
}
