using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Net.Sockets;
using System.IO;
using System.Net;
using System.Drawing.Printing;
using System.Drawing;
using Newtonsoft.Json.Converters;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using Newtonsoft.Json.Linq;
using System.Collections;

namespace myutils
{
    public class OfficeTools
    {
        static string temp = System.Environment.GetEnvironmentVariable("TEMP");
        public void printWord(string wordname)
        {

            // Create an Application object

            Word._Application wordApp = new Word.Application();
           
            object unknow = Type.Missing;
            object saveChanges = Microsoft.Office.Interop.Word.WdSaveOptions.wdSaveChanges;
            try
            {
              

                // I'm setting all of the alerts to be off as I am trying to get
                // this to print in the background
                wordApp.DisplayAlerts = Microsoft.Office.Interop.Word.WdAlertLevel.wdAlertsNone;
                // Open the document to print
                object filename = wordname;
                object missingValue = Type.Missing;
                // Using OpenOld so as to be compatible with other versions of Word
                Word._Document document = wordApp.Documents.OpenOld(
                    ref filename,
                    ref missingValue, ref missingValue,
                    ref missingValue, ref missingValue, ref missingValue,
                    ref missingValue, ref missingValue, ref missingValue, ref missingValue);

                // Set the active printer
                //wordApp.ActivePrinter = "My Printer Name";
                object myTrue = true; // Print in background
                object myFalse = false;
                // Using PrintOutOld to be version independent
                Console.WriteLine("准备开始打印...");
                wordApp.ActiveDocument.PrintOut(ref myTrue,
                    ref myFalse, ref missingValue, ref missingValue, ref missingValue,
                    ref missingValue, ref missingValue, ref missingValue, ref missingValue, ref missingValue, ref myFalse,
                    ref missingValue, ref missingValue, ref missingValue);
                Console.WriteLine("打印完成 ...");
                document.Close(ref missingValue, ref missingValue, ref missingValue);

                // Make sure all of the documents are gone from the queue   
                while (wordApp.BackgroundPrintingStatus > 0)
                {
                    System.Threading.Thread.Sleep(250);
                }
                //wordApp.Quit(ref saveChanges, ref unknow, ref unknow);
            }
            catch (Exception e)
            {
                Console.WriteLine(e.Message);

            }
            finally
            {
                
                if (wordApp != null)
                {
                    wordApp.Quit(ref saveChanges, ref unknow, ref unknow);
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(wordApp);
                    wordApp = null;
                }
                GC.Collect();
            }
           

        }

        public void printOffice(string ext,string filename)
        {
            switch (ext) {                
                case "doc":
                case "docx": 
                printWord(filename); 
                break;
                default: throw new FormatException("只支持Word文件格式");
                    
            }

        }       
 
        public void startServer(string host, Int32 port = 10086)
        {
            TcpListener server = null;
            try
            {
                // Set the TcpListener on port 13000.
                IPAddress localAddr = IPAddress.Parse(host);

                // TcpListener server = new TcpListener(port);
                server = new TcpListener(localAddr, port);

                // Start listening for client requests.
                server.Start();

                // Buffer for reading data
                Byte[] bytes = new Byte[256];


                // Enter the listening loop.
                while (true)
                {

                    Console.WriteLine("开始监听,等待用户...");
                    // Perform a blocking call to accept requests.
                    // You could also user server.AcceptSocket() here.
                    TcpClient client = server.AcceptTcpClient();
                    string name = System.Guid.NewGuid().ToString();
                    string fp = temp + @"\" + name;
                    if (File.Exists(fp))
                    {
                        File.Delete(fp);
                    }                    
                    // Get a stream object for reading and writing
                    NetworkStream stream = client.GetStream();
                    int i = 0;
                    using (FileStream fs = new FileStream(fp, FileMode.OpenOrCreate, FileAccess.Write, FileShare.Write))
                    {

                        // Loop to receive all the data sent by the client.
                        while ((i = stream.Read(bytes, 0, bytes.Length)) > 0)
                        {
                            // Translate data bytes to a ASCII string.                            
                            fs.Write(bytes, 0, i);
                            
                        }

                    }
                    Console.WriteLine("接受完毕,准备开始打印....");
                    printWord(fp);
                    client.Close();

                }
            }
            catch (SocketException e)
            {
                Console.WriteLine("SocketException: {0}", e);
            }
            finally
            {
                // Stop listening for new clients.
                server.Stop();
            }

        }

        /**
         * 
         * 连接到服务端,并发送文件 
         *       
         */
        public void Connect(string server, string fullpath)
        {
            try
            {
                Int32 port = 10086;
                TcpClient client = new TcpClient(server, port);


                // Get a client stream for reading and writing.
                //  Stream stream = client.GetStream();

                NetworkStream stream = client.GetStream();
                using (FileStream fs = File.OpenRead(fullpath))
                {
                    byte[] b = new byte[1024];
                    int i = 0;

                    while ((i = fs.Read(b, 0, b.Length)) > 0)
                    {
                        stream.Write(b, 0, i);
                    }

                }


                // Send the message to the connected TcpServer. 

                stream.Close();
                client.Close();
            }
            catch (ArgumentNullException e)
            {
                Console.WriteLine("{0}", e.Message);
            }
            catch (SocketException e)
            {
                Console.WriteLine("{0}", e.Message);
            }
        }           
         

        /**
         * 打印服务
         *          
         */
        public void startJsonServer(string host, Int32 port=10086)
        {

            TcpListener server = null;
            try
            {
                // Set the TcpListener on port 13000.
                IPAddress localAddr = IPAddress.Parse(host);

                // TcpListener server = new TcpListener(port);
                server = new TcpListener(localAddr, port);

                // Start listening for client requests.
                server.Start();

                // Buffer for reading data
                Byte[] bytes = new Byte[256];


                // Enter the listening loop.
                while (true)
                {

                    Console.WriteLine("数据打印服务正在监听,等待用户接入...");
                    // Perform a blocking call to accept requests.
                    // You could also user server.AcceptSocket() here.
                    TcpClient client = server.AcceptTcpClient();
                   
                    // Get a stream object for reading and writing
                    NetworkStream stream = client.GetStream();
                                        
                    int i = 0;

                    String jsonstr = "";

                   
                    using (MemoryStream fs = new MemoryStream())
                    {

                        // Loop to receive all the data sent by the client.
                        while ((i = stream.Read(bytes, 0, bytes.Length)) > 0)
                        {
                            // Translate data bytes to a ASCII string.                            
                            fs.Write(bytes, 0, i);  
                            
                        }

                        
                        byte[] byteArray = fs.ToArray();

                        jsonstr = System.Text.Encoding.Default.GetString(byteArray);
                       

                    }
                    client.Close();
                    

                    PdfContent pdf = JsonConvert.DeserializeObject<PdfContent>(jsonstr);
                    Console.WriteLine("输入文件是:"+pdf.pdfName);
                    Console.WriteLine("接受完毕,准备打印....");

                    string name = pdf.pdfName ;
                    string fp = temp + @"\" + name+"."+pdf.ext;
                    if (File.Exists(fp))
                    {
                        File.Delete(fp);
                    }

                    byte[] fbs = System.Convert.FromBase64String(pdf.data);
                    MemoryStream tms = new MemoryStream(fbs);
                    int j = 0;
                    using (FileStream tfs = new FileStream(fp, FileMode.CreateNew, FileAccess.ReadWrite, FileShare.ReadWrite))
                    {

                        // Loop to receive all the data sent by the client.
                        while ((j = tms.Read(bytes, 0, bytes.Length)) > 0)
                        {
                            // Translate data bytes to a ASCII string.                            
                            tfs.Write(bytes, 0, j);
                        }
                        tfs.Flush();
                        tfs.Close();

                    }                    
                    
                    printOffice(pdf.ext, fp);
                                      

                }
            }
            catch (SocketException e)
            {
                Console.WriteLine("SocketException: {0}", e);
            }
            finally
            {
                // Stop listening for new clients.
                server.Stop();
            }
        }

       

        public byte[] StreamToBytes(Stream stream) 
        { 
            
            byte[] bytes = new byte[stream.Length]; 
            stream.Read(bytes, 0, bytes.Length); 
            // 设置当前流的位置为流的开始 
            stream.Seek(0, SeekOrigin.Begin); 
            return bytes; 
        }


    }

    class PdfContent
    {

        public string pdfName { get; set; }

        public string data { get; set; }

        public string ext { get; set; }


    }
    public class ComboxItem
    {
        public string text { get; set; }
        public string value { get; set; }
    }
}