using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Timers;
using Newtonsoft.Json;

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
        private string _topic;

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
            return 1;
        }

        private void job(object sender, EventArgs args) { _callback.UpdateNotify(); }

        // Called when Excel no longer requires RTD topics from the RealTimeData server
        public void ServerTerminate()
        {
            if (_timer != null)
            {
                _timer.Dispose();
                _timer = null;
            }
        }

        // Called by Excel if a given interval has elapsed since the last time
        // Excel was notified of updates from the RealTimeData server
        public int Heartbeat() { return 1; }

        // Called whenever Excel requests a new RTD topic from the RealTimeData server
        public dynamic ConnectData(int topicId, ref Array strings, ref bool GetNewValues)
        {
            Console.WriteLine("ConnectData topicId: {0}, Array: {1}", topicId, JsonConvert.SerializeObject(strings));
            _topic = strings.GetValue(0).ToString();
            return getData(_topic);
        }

        // Called whenever Excel no longer requires a specific topic
        public void DisconnectData(int topicId) { }

        // Called when Excel is requesting a refresh on topics.
        // RefreshData will be called after an UpdateNotify has been issued by the server.
        // This event should:
        // - supply a value for TopicCount (number of topics to update)
        // - return a two dimensional variant array containing the topic ids and the new values of each.
        public Array RefreshData(ref int TopicCount)
        {
            object[,] data = new object[10, 10];
            for (int i = 0; i < 4; i++)
            {
                data[0, i] = i;
                data[1, i] = getData(_topic);
                TopicCount = i + 1;
            }
            _timer.Start();
            return data;
        }

        private string getData(string symbol)
        {
            try { return "data from hello_rtd_excel_0: " + _rng.Next(0, 100).ToString(); }
            catch (Exception ex) { return "error when getData: " + ex.Message; }
        }

        private string now() { return DateTime.UtcNow.ToString("o") + ": "; }
    }
}
