using Microsoft.Office.Interop.Excel;
using System;
using System.Runtime.InteropServices;
using System.Timers;

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
        private string _symbol;

        public RTDDemoServer()
        {
            _rng = new Random();
        }

        public int ServerStart(IRTDUpdateEvent CallbackObject)
        {
            _callback = CallbackObject;
            _timer = new Timer();
            _timer.Elapsed += new ElapsedEventHandler(job);
            _timer.Interval = 1000;
            return 1;
        }

        private void job(object sender, EventArgs args) { _callback.UpdateNotify(); }

        public void ServerTerminate()
        {
            if (_timer != null)
            {
                _timer.Dispose();
                _timer = null;
            }
        }

        public int Heartbeat() { return 1; }

        public dynamic ConnectData(int topicId, ref Array strings, ref bool GetNewValues)
        {
            _symbol = strings.GetValue(0).ToString();
            _timer.Start();
            return getData(_symbol);
        }

        public void DisconnectData(int topicId) { }

        public Array RefreshData(ref int TopicCount)
        {
            object[,] data = new object[10, 10];
            for (int i = 0; i < 4; i++)
            {
                data[0, i] = i;
                data[1, i] = getData(_symbol);
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
    }
}
