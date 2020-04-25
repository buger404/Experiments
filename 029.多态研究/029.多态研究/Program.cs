using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace _029.多态研究
{
    public class GrandPa
    {
        public virtual void Shit()
        {

        }
    }
    public class Dad : GrandPa
    {
        public override void Shit()
        {
            Console.WriteLine("so this is your dear dad~");
        }
    }
    public class Son : Dad
    {
        public override void Shit()
        {

        }
    }
    public class MagicRandom : Random
    {
        //public override int Next()
        //{
        //    return 404233;
        //}
        public new int Next()
        {
            return 23333;
        }
        public override string ToString()
        {
            return "不许ToString！";
        }
    }
    class Program
    {
        static void Main(string[] args)
        {
            Random r = new Random();
            Console.WriteLine(r.Next());
            Console.WriteLine(r.ToString());
            Console.WriteLine("======================");
            MagicRandom mr = new MagicRandom();
            Console.WriteLine(mr.Next());
            Console.WriteLine(mr.ToString());
            Console.WriteLine("======================");
            Random rmr = new MagicRandom();
            Console.WriteLine(rmr.Next());
            Console.WriteLine(rmr.ToString());
            Console.WriteLine("======================");
            Console.ReadLine();
        }
    }
}
