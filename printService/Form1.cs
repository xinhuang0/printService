using Fleck;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using printService.printModel;
using Seagull.BarTender.Print;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace printService
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            Dosocket();

        }

        /// <summary>
        /// 开启websocket服务器
        /// </summary>
        public bool Dosocket()
        {
            bool result = false;
            try
            {
                //存放连接服务器的socket对象
                var allScokets = new List<IWebSocketConnection>();
                var server = new WebSocketServer("ws://10.205.172.141:5555");    //创建webscoket服务端实例
                server.Start(socket =>
                {
                    //打开链接
                    socket.OnOpen = () =>
                    {
                        //Console.WriteLine("Open");
                        allScokets.Add(socket);
                        result = true;
                    };
                    //关闭链接;
                    socket.OnClose = () =>
                    {
                        //Console.WriteLine("Close");
                        allScokets.Remove(socket);
                    };
                    //监听发送数据
                    socket.OnMessage = message =>
                    {
                        //Console.WriteLine(message);
                        if (!string.IsNullOrEmpty(message))
                        {
                            AjaxResult res = new AjaxResult
                            {
                                Success = true,
                                Msg = "请求失败！",
                                Data = null
                            };
                            dynamic objdatasTem = JObject.Parse(message);
                            var objdatas = objdatasTem.Msg;
                            if (objdatas == "请求链接")
                            {
                                //PrintDataSocketTest();
                                res.Msg = "链接成功";
                                socket.Send(JsonConvert.SerializeObject( res));
                            }
                            else
                            {
                                
                                //var print = PrintDataSocket(objdatasTem.photo, JsonConvert.SerializeObject(objdatasTem.Data[0]));
                                var print = PrintDataSocket(Convert.ToString( objdatasTem.photo), JsonConvert.SerializeObject(objdatasTem.Data[0]));
                                if (print)
                                {
                                    res.Msg = "打印成功";
                                    socket.Send(JsonConvert.SerializeObject(res));
                                }
                                else
                                {
                                    res.Msg = "打印失败";
                                    socket.Send(JsonConvert.SerializeObject(res));
                                }
                            }

                        }

                    };
                });
            }
            catch (Exception ex)
            {

                result = false;
            }
            return result;
            //while (input != "exit")
            //{
            //    foreach (var socket in allScokets.ToList())
            //    {
            //        socket.Send("服务端：" + input);
            //    }
            //    input = Console.ReadLine();
            //}

        }
        public bool PrintDataSocketTest()
        {
            bool result = false;

            try

            {

                Engine btEngine = new Engine();

                btEngine.Start();

                string lj = AppDomain.CurrentDomain.BaseDirectory + "访客标签.btw";  //test.btw是BT的模板
                //string lj = Server.MapPath("~/printTemp/访客标签.btw");  //test.btw是BT的模板
                LabelFormatDocument btFormat = btEngine.Documents.Open(lj);

                //对BTW模版相应字段进行赋值 

                //btFormat.SubStrings["name"].Value = "";//访问人姓名
                //btFormat.SubStrings["name"].Value = objdatas?.lfMainName;//访问人姓名
                //btFormat.SubStrings["visitIDnum"].Value = objdatas?.lfIdCard;//身份证
                //btFormat.SubStrings["visit_Sdate"].Value = objdatas?.lfStartTime;//来访开始时间
                //btFormat.SubStrings["visit_Edate"].Value = objdatas?.lfEndTime;//来访结束时间
                //btFormat.SubStrings["visiterPhone"].Value = objdatas?.status;//来访者电话
                //btFormat.SubStrings["InterV_dep"].Value = objdatas?.lfMainName;//被访人部门
                //btFormat.SubStrings["Interviewee"].Value = objdatas?.applyPerson;//被访人
                //btFormat.SubStrings["InterV_phone"].Value = objdatas?.lfMainName;//被访问人电话
                //btFormat.SubStrings["visitReason"].Value = objdatas?.lfAims;//事由
                //btFormat.SubStrings["visiterAre"].Value = objdatas?.lfactive;//访问区域

                //指定打印机名 

                btFormat.PrintSetup.PrinterName = @"\\10.205.173.44\EPSON TM-C3520 Ver2";

                //改变标签打印数份连载 

                btFormat.PrintSetup.NumberOfSerializedLabels = 1;

                //打印份数                   

                btFormat.PrintSetup.IdenticalCopiesOfLabel = 1;

                Messages messages;

                int waitout = 10000; // 10秒 超时 

                Result nResult1 = btFormat.Print("标签打印软件", waitout, out messages);

                btFormat.PrintSetup.Cache.FlushInterval = CacheFlushInterval.PerSession;

                //不保存对打开模板的修改 

                btFormat.Close(SaveOptions.DoNotSaveChanges);

                //结束打印引擎                  

                btEngine.Stop();
                result = true;
            }

            catch (Exception ex)
            {

                result = false;

            }




            return result;
        }

        public bool PrintDataSocket(string phonestr, string oainfo)
        {
            bool result = false;

            try
            {
                if (!string.IsNullOrEmpty(oainfo))
                {

                    string printTemp = "访客标签.btw";
                    dynamic objdatasTem = JObject.Parse(oainfo);
                    var objdatas = objdatasTem;
                    string areaTemp = objdatas?.lfactive;
                    printTemp = CheckVisitArea(areaTemp) =="red"? "访客标签.btw" : "访客标签绿色通行.btw";
                    Engine btEngine = new Engine();
                    btEngine.Start();

                    string lj = AppDomain.CurrentDomain.BaseDirectory + printTemp;  //test.btw是BT的模板
                                                                                       //string lj = Server.MapPath("~/printTemp/访客标签.btw");  //test.btw是BT的模板
                    LabelFormatDocument btFormat = btEngine.Documents.Open(lj);

                    //对BTW模版相应字段进行赋值 
                    btFormat.SubStrings["headPhotos"].Value = phonestr;//访问人姓名
                    btFormat.SubStrings["name"].Value = objdatas?.lfMainName;//访问人姓名
                    btFormat.SubStrings["name"].Value = objdatas?.lfMainName;//访问人姓名
                    btFormat.SubStrings["visitIDnum"].Value = objdatas?.lfIdCard;//身份证
                    btFormat.SubStrings["visit_Sdate"].Value = objdatas?.lfStartTime;//来访开始时间
                    btFormat.SubStrings["visit_Edate"].Value = objdatas?.lfEndTime;//来访结束时间
                    btFormat.SubStrings["visiterPhone"].Value = objdatas?.status;//来访者电话
                    btFormat.SubStrings["InterV_dep"].Value = objdatas?.lfMainName;//被访人部门
                    btFormat.SubStrings["Interviewee"].Value = objdatas?.applyPerson;//被访人
                    btFormat.SubStrings["InterV_phone"].Value = objdatas?.lfMainName;//被访问人电话
                    btFormat.SubStrings["visitReason"].Value = objdatas?.lfAims;//事由
                    btFormat.SubStrings["visiterAre"].Value = objdatas?.lfactive;//访问区域

                    //指定打印机名 

                    btFormat.PrintSetup.PrinterName = @"\\10.205.173.44\EPSON TM-C3520 Ver2";

                    //改变标签打印数份连载 

                    btFormat.PrintSetup.NumberOfSerializedLabels = 1;

                    //打印份数                   

                    btFormat.PrintSetup.IdenticalCopiesOfLabel = 1;

                    Messages messages;

                    int waitout = 10000; // 10秒 超时 

                    Result nResult1 = btFormat.Print("标签打印软件", waitout, out messages);

                    btFormat.PrintSetup.Cache.FlushInterval = CacheFlushInterval.PerSession;

                    //不保存对打开模板的修改 

                    btFormat.Close(SaveOptions.DoNotSaveChanges);

                    //结束打印引擎                  

                    btEngine.Stop();
                    result = true;
                }
            }

            catch (Exception ex)

            {

                result = false;



            }




            return result;
        }
        private string CheckVisitArea(string visitarea)
        {
            string result = "green";

            var areaTemp = visitarea?.Split(';');

            for (int i = 0; i < areaTemp.Length; i++)
            {

                if (areaTemp[i].Contains("车间") || areaTemp[i].Contains("17栋3楼办公区域"))
                {
                    result = "red";
                }
                //else if (areaTemp[i].Contains("办公区域") && areaTemp[i] != "17栋3楼办公区域")
                //{
                //    result = "background:green";
                //}
                //else
                //{
                //    result = "background:green";
                //}
            }

            return result;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            PrintDataSocketTest();
        }
    }
}
