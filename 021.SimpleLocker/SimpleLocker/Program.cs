using System;
using System.IO;

namespace SimpleLocker
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.ForegroundColor = ConsoleColor.Green;
            Console.WriteLine("Error 404 垃圾文件加密解密工具\n版本号：beta228.01");
            string file;
            if (args.Length == 0)
            {
                Console.Write("目标文件：");
                file = Console.ReadLine();
            }
            else
            {
                file = args[0];
                Console.WriteLine("目标文件：" + file);
            }
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("请提供加密密钥：");
            string key = Console.ReadLine();
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write("现在开始加密？[Y-加密/N-解密]：");
            ConsoleKeyInfo b = Console.ReadKey();
            byte[] fbyte;int fpos = 0;int locks;int step = 0;
            try
            {
                if (b.Key == ConsoleKey.Y)
                {
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("\n开始执行加密操作...");
                    fbyte = File.ReadAllBytes(file);
                    File.Copy(file, file + "[备份]");
                    Console.WriteLine("备份已创建！");
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.Write("加密：[");
                    for (int i = 0;i < fbyte.Length; i++)
                    {
                        if((int)(i / fbyte.Length * 100) / 5 > step)
                        {
                            Console.Write("█");step++;
                        }
                        locks = (((int)fbyte[i] + key[fpos] * key.Length) % 255);
                        fbyte[i] = (byte)locks;
                        fpos++;
                        if (fpos >= key.Length) fpos = 0;
                    }
                    for(int j = step;j < 20; j++)
                    {
                        Console.Write("█");
                    }
                    Console.Write("]"); 
                    File.WriteAllBytes(file, fbyte);
                    Console.ForegroundColor = ConsoleColor.Yellow;
                    Console.WriteLine("\n加密完成！");
                }
                else
                {
                    if (b.Key == ConsoleKey.N)
                    {
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine("\n开始执行解密操作...");
                        fbyte = File.ReadAllBytes(file);
                        File.Copy(file, file + "[备份]");
                        Console.WriteLine("备份已创建！");
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.Write("解密：[");
                        for (int i = 0; i < fbyte.Length; i++)
                        {
                            if ((int)(i / fbyte.Length * 100) / 5 > step)
                            {
                                Console.Write("█"); step++;
                            }
                            locks = Math.Abs(((int)fbyte[i] - key[fpos] * key.Length)) % 255;
                            fbyte[i] = (byte)locks;
                            fpos++;
                            if (fpos >= key.Length) fpos = 0;
                        }
                        for (int j = step; j < 20; j++)
                        {
                            Console.Write("█");
                        }
                        Console.Write("]");
                        File.WriteAllBytes(file, fbyte);
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine("\n解密完成！");
                    }
                    else
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("\n无效操作。");
                    }
                }
            }
            catch(Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("\n操作异常：" + e.Message);
            }

            Console.ReadLine();
        }
    }
}
