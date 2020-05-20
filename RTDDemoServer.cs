using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Timers;
using System.Collections.Generic;

using Microsoft.Office.Interop.Excel;
using Newtonsoft.Json;
using WebSocketSharp;

namespace RtdServer
{
    [
        ComVisible(true),
        Guid("dcc6eaaf-f6c9-46f3-8cb5-bfab118e6099"),
        ProgId("hello_rtd_excel"),
    ]
    public class RTDDemoServer : IRtdServer
    {
        private IRTDUpdateEvent _callback;
        private Timer _timer;
        private Random _rng;
        private Dictionary<int, string> _topics; // map topicId to topicName
        private WebSocket _wsConn;
        private Dictionary<string, Security> _securitiesBoard;
        private bool _isElectronStopped;

        public RTDDemoServer()
        {
            FileStream _loggerP0 = new FileStream(
                @"C:\Users\Admin\Desktop\aaa.txt", FileMode.Create, FileAccess.Write);
            StreamWriter _loggerP1 = new StreamWriter(_loggerP0);
            _loggerP1.AutoFlush = true;
            Console.SetOut(_loggerP1); Console.SetError(_loggerP1);
            Console.WriteLine(now() + "hello_rtd_excel RTDDemoServer constructor");
        }

        // Called when Excel requests the first RTD topic for the server. 
        // ServerStart should return a 1 on success, 0 on failure. 
        // The first parameter of the ServerStart method is a callback object 
        // that the RealTimeData server uses to notify Excel when it should 
        // gather updates from the RealTimeData server
        public int ServerStart(IRTDUpdateEvent CallbackObject)
        {
            _callback = CallbackObject;
            _timer = new Timer();
            _timer.Elapsed += new ElapsedEventHandler(job);
            _timer.Interval = 1000;
            _timer.Start();
            _rng = new Random();
            _topics = new Dictionary<int, string>();
            _securitiesBoard = new Dictionary<string, Security>();
            _isElectronStopped = false;

            try
            {
                var wsServerAddr = "ws://10.100.50.102:8001/";
                var ws = new WebSocket(wsServerAddr);
                ws.OnMessage += (sender, e) =>
                {
                    Console.WriteLine("{0} received msg: {1}", now(), e.Data);
                    var xMsg = JsonConvert.DeserializeObject<XSecurity>(e.Data);
                    var u = xMsg.data;
                    var sec = new Security();
                    if (_securitiesBoard.ContainsKey(u.code))
                    {
                        sec = _securitiesBoard[u.code];
                        if (u.referencePrice != 0) { sec.referencePrice = u.referencePrice; }
                        if (u.last != 0) { sec.last = u.last; sec.change = u.change; }
                        sec.bidPrice = u.bidPrice;
                        sec.bidVolume = u.bidVolume;
                    }
                    else
                    {
                        sec = u;
                    }
                    _securitiesBoard[u.code] = sec;
                };
                ws.Connect();
                Console.WriteLine("connected to {0}", wsServerAddr);
            }
            catch (Exception err)
            {
                Console.WriteLine("error when connect WebSocket: 0}", err);
            }
            return 1;
        }

        private void job(object sender, EventArgs args) { 
            _callback.UpdateNotify();
            var techXApp = "daominah_electron_demo";
            System.Diagnostics.Process[] pname = 
                System.Diagnostics.Process.GetProcessesByName(techXApp);
            if (pname.Length == 0) {
                _isElectronStopped = true;
                // Console.WriteLine("{0} stopped", techXApp);
            }
            else { 
                _isElectronStopped = false;
                // Console.WriteLine("{0} is running", techXApp);
            }
        }

        // Called by Excel if a given interval has elapsed since the last time
        // Excel was notified of updates from the RealTimeData server
        public int Heartbeat() { return 1; }

        // Called whenever Excel no longer requires a specific topic
        public void DisconnectData(int topicId) { }

        // Called when Excel no longer requires RTD topics from the RealTimeData server
        public void ServerTerminate()
        {
            if (_timer != null)
            {
                _timer.Dispose();
                _timer = null;
            }
        }

        // Called whenever Excel requests a new RTD topic from the RealTimeData server
        public dynamic ConnectData(int topicId, ref Array strings, ref bool GetNewValues)
        {
            Console.WriteLine("{0} ConnectData topicId: {1}, strings: {2}",
                now(), topicId, json(strings));
            string topicName = strings.GetValue(0).ToString();
            _topics[topicId] = topicName;
            Console.WriteLine("_topics: {0}", json(_topics));
            return getData(topicName);
        }

        // Called when Excel is requesting a refresh on topics.
        // RefreshData will be called after an UpdateNotify has been issued by the server.
        // This event should:
        // - supply a value for TopicCount (number of topics to update)
        // - return a two dimensional variant array containing the topic ids and the new values of each.
        public Array RefreshData(ref int TopicCount)
        {
            object[,] data = new object[_topics.Count, _topics.Count];
            TopicCount = _topics.Count;
            for (int i = 1; i <= _topics.Count; i++)
            {
                data[0, i - 1] = i;
                data[1, i - 1] = getData(_topics[i]);
            }
            // Console.WriteLine("{0} RefreshData: {1}", now(), json(data));
            return data;
        }

        private string getData(string topic)
        {
            if (_isElectronStopped) return "plz turn on TechX app";
            try
            {
                return json(_securitiesBoard[topic]);
            }
            catch (Exception err)
            {
                // Console.WriteLine("error when getData: {0}", err);
                return String.Format("{0}'s data is not available", topic);
            }
        }

        private string now() { return DateTime.UtcNow.ToString("o") + ": "; }
        private string json(object obj) { return JsonConvert.SerializeObject(obj); }
    }

    public class Security
    {
        public string code;
        public float referencePrice;
        public float last;
        public float change;
        public float bidPrice;
        public float bidVolume;
    }

    public class XSecurity
    {
        public string sourceId;
        public Security data;
    }
}
