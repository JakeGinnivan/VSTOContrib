using System;
using System.Diagnostics;
using System.Windows.Forms;

namespace FacebookToOutlookAddin.Demo
{
    public class MessageBoxTimer : IDisposable
    {
        private readonly TickTimer _timer;
        public MessageBoxTimer(string message)
        {
            Message = message;
            _timer = new TickTimer();
        }

        public string Message { get; set; }

        public void Dispose()
        {
            _timer.Stop();
            MessageBox.Show(string.Format("{0}: Total {1:0.0000} ms\r\n ", Message, _timer.ElapsedMilliseconds));
            Debug.WriteLine(string.Format("{0}: Total {1:0.0000} ms\r\n ", Message, _timer.ElapsedMilliseconds));
        }
    }
}
