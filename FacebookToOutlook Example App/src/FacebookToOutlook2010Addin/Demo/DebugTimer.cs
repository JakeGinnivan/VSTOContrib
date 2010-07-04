using System;
using System.Diagnostics;

namespace FacebookToOutlookAddin.Demo
{
    public class DebugTimer : IDisposable
    {
        private readonly string _message;
        private readonly TickTimer _timer;
        public DebugTimer(string message)
        {
            _message = message;
            _timer = new TickTimer();
        }

        public void Dispose()
        {
            _timer.Stop();
            Trace.WriteLine(string.Format("{0}: Total {1:0.0000} ms\r\n ", _message, _timer.ElapsedMilliseconds));
        }
    }
}
