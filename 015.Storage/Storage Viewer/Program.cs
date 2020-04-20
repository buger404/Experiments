using DataArrange.Storages;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using static DataArrange.Storages.Storage;

namespace Storage_Viewer
{
    class Program
    {
        static void help()
        {
            Console.ForegroundColor = ConsoleColor.Gray;
            Console.WriteLine("Storage Viewer , ver 0.0.5");
            Console.WriteLine("#GUIDENCE");
            Console.WriteLine("ls : list all data");
            Console.WriteLine("lsu <user> : list someone's data");
            Console.WriteLine("lsui <userindex> : list someone's data");
            Console.WriteLine("lsus : list all users");
            Console.WriteLine("sek <content> : search all data");
            Console.WriteLine("edt <user> <item> <content> : edit data");
            Console.WriteLine("get <user> <item> : show data");
            Console.WriteLine("geti <userindex> <item> : show data");
            Console.WriteLine("back : create a backup file");
            Console.WriteLine("out : close file and open another");
            Console.WriteLine("help : show guidence");
        }
        static void Main(string[] args)
        {
            help();

        login:
            Console.ForegroundColor = ConsoleColor.Yellow;
            Console.Write("\nInput your save name : ");
            string name = Console.ReadLine();
            Storage save = new Storage(name);
            Console.WriteLine("");
            if(!File.Exists("C:\\DataArrange\\" + name + "-userdata.json"))
            {
                Console.ForegroundColor = ConsoleColor.Red;
                Console.WriteLine("Warning : this file does not exist, the tool will create a new file.\n");
            }

        appstart:
            Console.ForegroundColor = ConsoleColor.White;
            Console.Write(name + ">");
            string[] cmd = Console.ReadLine().Split(' ');
            DataArea daa;
            int index = 0;
            Console.WriteLine("");
            try
            {
                switch (cmd[0])
                {
                    case ("back"):
                        string backid = new Guid().GetHashCode().ToString();
                        File.Copy("C:\\DataArrange\\" + name + "-userdata.json",
                                  "C:\\DataArrange\\[" + backid + "]" + name + "-userdata.json");
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Succeed : " + "[" + backid + "]" + name + "-userdata.json");
                        break;
                    case ("edt"):
                        save.putkey(cmd[1], cmd[2], cmd[3]);
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("Succeed");
                        break;
                    case ("geti"):
                        daa = save.data.Areas[int.Parse(cmd[1])];
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("user(" + index + "): " + daa.User);
                        foreach (DataItem di in daa.Items.FindAll(m => m.Key == cmd[2]))
                        {
                            Console.ForegroundColor = ConsoleColor.DarkGreen;
                            Console.Write("     " + di.Key + " : ");
                            Console.ForegroundColor = ConsoleColor.Gray;
                            Console.Write(di.Value + "\n");
                        }
                        break;
                    case ("get"):
                        foreach (DataArea da in save.data.Areas.FindAll(m => m.User == cmd[1]))
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("user(" + index + "): " + da.User);
                            foreach (DataItem di in da.Items.FindAll(m => m.Key == cmd[2]))
                            {
                                Console.ForegroundColor = ConsoleColor.DarkGreen;
                                Console.Write("     " + di.Key + " : ");
                                Console.ForegroundColor = ConsoleColor.Gray;
                                Console.Write(di.Value + "\n");
                            }
                            index++;
                        }
                        break;
                    case ("out"): goto login;
                    case ("ls"):
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("owner:" + save.data.Owner);
                        foreach (DataArea da in save.data.Areas)
                        {
                            Console.ForegroundColor = ConsoleColor.Green;
                            Console.WriteLine("user(" + index + "): " + da.User);
                            foreach (DataItem di in da.Items)
                            {
                                Console.ForegroundColor = ConsoleColor.DarkGreen;
                                Console.Write("     " + di.Key + " : ");
                                Console.ForegroundColor = ConsoleColor.Gray;
                                Console.Write(di.Value + "\n");
                            }
                            index++;
                        }
                        break;
                    case ("lsu"):
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("owner:" + save.data.Owner);
                        foreach (DataArea da in save.data.Areas)
                        {
                            if(da.User == cmd[1])
                            {
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.WriteLine("user(" + index + "): " + da.User);
                                foreach (DataItem di in da.Items)
                                {
                                    Console.ForegroundColor = ConsoleColor.DarkGreen;
                                    Console.Write("     " + di.Key + " : ");
                                    Console.ForegroundColor = ConsoleColor.Gray;
                                    Console.Write(di.Value + "\n");
                                }
                            }
                            index++;
                        }
                        break;
                    case ("lsui"):
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("owner:" + save.data.Owner);
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("user(" + index + "): " + save.data.Areas[int.Parse(cmd[1])].User);
                        foreach (DataItem di in save.data.Areas[int.Parse(cmd[1])].Items)
                        {
                            Console.ForegroundColor = ConsoleColor.DarkGreen;
                            Console.Write("     " + di.Key + " : ");
                            Console.ForegroundColor = ConsoleColor.Gray;
                            Console.Write(di.Value + "\n");
                        }
                        break;
                    case ("lsus"):
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("owner:" + save.data.Owner);
                        foreach (DataArea da in save.data.Areas)
                        {
                            Console.ForegroundColor = ConsoleColor.White;
                            Console.WriteLine("user(" + index + "): " + da.User);
                            index++;
                        }
                        break;
                    case ("sek"):
                        Console.ForegroundColor = ConsoleColor.Green;
                        Console.WriteLine("owner:" + save.data.Owner);
                        cmd[1] = cmd[1].ToLower();
                        foreach (DataArea da in save.data.Areas)
                        {
                            if(da.User.ToLower().IndexOf(cmd[1]) >= 0)
                            {
                                Console.ForegroundColor = ConsoleColor.Gray;
                                Console.Write("user(" + index + "): " +
                                                    da.User.Substring(0, da.User.ToLower().IndexOf(cmd[1]))
                                                 );
                                Console.ForegroundColor = ConsoleColor.Green;
                                Console.Write(da.User.Substring(da.User.ToLower().IndexOf(cmd[1]),cmd[1].Length));
                                Console.ForegroundColor = ConsoleColor.Gray;
                                Console.Write(
                                    da.User.Substring(da.User.ToLower().IndexOf(cmd[1]) + cmd[1].Length)
                                    );
                                Console.Write("\n");
                            }
                            else
                            {
                                Console.ForegroundColor = ConsoleColor.Gray;
                                Console.WriteLine("user(" + index + "): " + da.User);
                            }

                            foreach (DataItem di in da.Items)
                            {
                                if (di.Key.ToLower().IndexOf(cmd[1]) >= 0 || di.Value.ToLower().IndexOf(cmd[1]) >= 0)
                                {
                                    Console.Write("     ");
                                    if (di.Key.ToLower().IndexOf(cmd[1]) >= 0)
                                    {
                                        Console.ForegroundColor = ConsoleColor.Gray;
                                        Console.Write(
                                                        di.Key.Substring(0, di.Key.ToLower().IndexOf(cmd[1]))
                                                     );
                                        Console.ForegroundColor = ConsoleColor.Green;
                                        Console.Write(di.Key.Substring(di.Key.ToLower().IndexOf(cmd[1]), cmd[1].Length));
                                        Console.ForegroundColor = ConsoleColor.Gray;
                                        Console.Write(
                                            di.Key.Substring(di.Key.ToLower().IndexOf(cmd[1]) + cmd[1].Length)
                                            );
                                    }
                                    else
                                    {
                                        Console.ForegroundColor = ConsoleColor.Gray;
                                        Console.Write(di.Key);
                                    }

                                    Console.ForegroundColor = ConsoleColor.Gray;
                                    Console.Write(" : ");
                                    if (di.Value.ToLower().IndexOf(cmd[1]) >= 0)
                                    {
                                        Console.ForegroundColor = ConsoleColor.Gray;
                                        Console.Write(
                                                        di.Value.Substring(0, di.Value.ToLower().IndexOf(cmd[1]))
                                                     );
                                        Console.ForegroundColor = ConsoleColor.Green;
                                        Console.Write(di.Value.Substring(di.Value.ToLower().IndexOf(cmd[1]), cmd[1].Length));
                                        Console.ForegroundColor = ConsoleColor.Gray;
                                        Console.Write(
                                            di.Value.Substring(di.Value.ToLower().IndexOf(cmd[1]) + cmd[1].Length)
                                            );
                                    }
                                    else
                                    {
                                        Console.ForegroundColor = ConsoleColor.Gray;
                                        Console.Write(di.Value);
                                    }

                                    Console.Write("\n");
                                }
                            }
                            
                            index++;
                        }
                        break;
                    case ("help"): help(); break;
                    default: Console.ForegroundColor = ConsoleColor.Red; Console.WriteLine("unknown command '" + cmd[0] + "'"); break;
                }
            }
            catch(Exception e)
            {
                Console.ForegroundColor = ConsoleColor.Red; 
                Console.WriteLine(e.Message); 
            }
            Console.WriteLine("");
            goto appstart;
        }
    }
}
