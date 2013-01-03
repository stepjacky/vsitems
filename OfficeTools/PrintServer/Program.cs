using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using myutils;
using System.Net;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using Newtonsoft.Json.Linq;
using System.Text.RegularExpressions;
using System.Threading;
namespace PrintServer
{
    class Program
    {
        static OfficeTools tools = new OfficeTools();
        static string ipaddr = "localhost";
        static void Main(string[] args)
        {

            string ov = System.Environment.OSVersion.Version.Major.ToString();
            Console.WriteLine(ov);

            IPHostEntry ipHost = Dns.GetHostEntry(Dns.GetHostName());
            int i = 0;
            string[] iplist = new string[ipHost.AddressList.Count()];
            
            foreach (IPAddress ip in ipHost.AddressList)
            {
                int j = i++;
                iplist[j] = ip.ToString(); 
                Console.WriteLine(j+"."+ ip.ToString());
            }

            Console.WriteLine("输入IP地址前面的数字,选择对应的IP监听,输入q 退出程序");
            Console.WriteLine("请选择一个IPV4地址:");
           
            bool hasListening = false;
            Thread uithread = new Thread(new ThreadStart(startDataListener));
            uithread.IsBackground = true;
            while (true)
            {

                try
                {
                    string rl = Console.ReadLine();
                    if ("q" == rl || "Q" == rl)
                    {
                       
                        if (uithread != null && uithread.IsAlive)
                        {
                            Console.WriteLine("三秒后自动退出...");
                            uithread.Abort();
                            uithread.Join(3000);
                        }
                        Environment.Exit(0);
                    }
                    else
                    {
                        if (hasListening)
                            Console.WriteLine("数据服务已经启动,输入Q退出");
                       
                    }
                    if (!hasListening)
                    {
                        i = Int32.Parse(rl);
                        string fPattern = @"((2[0-4]\d|25[0-5]|[01]?\d\d?)\.){3}(2[0-4]\d|25[0-5]|[01]?\d\d?)";
                        String ip = Regex.Match(iplist[i], fPattern).Value;
                        if (null == ip || "" == ip)
                        {
                            Console.WriteLine("请选择一个IPV4的地址");
                            continue;
                        }
                        Console.WriteLine("您选择的IP是:" + ip);
                        ipaddr = ip;
                        uithread.Start();
                        hasListening = true;
                    }
                }
                catch (Exception fe)
                {
                    Console.WriteLine(fe.Message+" ,输入Q退出");
                    continue;

                }
               
            }
        }
        static void startDataListener()
        {
            tools.startJsonServer(ipaddr);
        }
        
    }   
}
