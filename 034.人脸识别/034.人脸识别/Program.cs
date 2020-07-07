using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Runtime.Serialization;
using Newtonsoft.Json.Linq;

namespace _034.人脸识别
{
    class Program
    {
        public static string GetRespondText(string host,string method = "post",string extraArgs = "")
        {
            //创建请求
            HttpWebRequest re = (HttpWebRequest)WebRequest.Create(host);
            re.Method = method; re.KeepAlive = true;

            //填写表单
            if(extraArgs != "")
            {
                byte[] buffer = Encoding.Default.GetBytes(extraArgs);
                re.ContentLength = buffer.Length;
                re.GetRequestStream().Write(buffer, 0, buffer.Length);
            }

            //取响应
            HttpWebResponse ret = (HttpWebResponse)re.GetResponse();
            string result = new StreamReader(ret.GetResponseStream(), Encoding.Default).ReadToEnd();

            return result;
        }
        static int Main(string[] args)
        {
            string srcimg, tarimg;
            if(args.Length == 0)
            {
                Console.Write("Source face: ");
                srcimg = Console.ReadLine();
                Console.Write("Target face: ");
                tarimg = Console.ReadLine();
            }
            else
            {
                srcimg = args[0]; tarimg = args[1];
            }
            
            srcimg = Convert.ToBase64String(File.ReadAllBytes(srcimg));
            tarimg = Convert.ToBase64String(File.ReadAllBytes(tarimg));

            string id = "mRSQCrggiSov22o6EXUCZRlV";
            string secret = "sIHnGMAXa4gFQOlZSAoGkbaUGS1uC9kV";

            string res;

            res = GetRespondText($"https://aip.baidubce.com/oauth/2.0/token?grant_type=client_credentials&client_id={id}&client_secret={secret}&","get");

            JObject tokenret = JObject.Parse(res);

            string token = tokenret["access_token"].ToString();

            Console.WriteLine($"Connected, checking...");

            string arg = "[{\"image\": \"{src}\", \"image_type\": \"BASE64\", \"face_type\": \"LIVE\", \"quality_control\": \"LOW\"},{\"image\": \"{tar}\", \"image_type\": \"BASE64\", \"face_type\": \"LIVE\", \"quality_control\": \"LOW\"}]";
            arg = arg.Replace("{src}", srcimg).Replace("{tar}", tarimg);
            res = GetRespondText($"https://aip.baidubce.com/rest/2.0/face/v3/match?access_token={token}","post",arg);

            JObject faceret = JObject.Parse(res);

            string err = faceret["error_msg"].ToString();

            if(err != "SUCCESS")
            {
                Console.WriteLine($"Error: {err}");
                if(args.Length > 0) return 0;
            }
            else
            {
                Console.WriteLine($"Match: {faceret["result"]["score"]}");
                if (args.Length > 0) return Convert.ToInt32(float.Parse(faceret["result"]["score"].ToString()));
            }
            Console.ReadLine();
            return 0;
        }
    }
}
