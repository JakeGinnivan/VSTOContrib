using System;
using System.Runtime.InteropServices;

namespace FacebookToOutlookAddin.Demo
{
    /// <summary>
    /// A high resolution time that also records start time.
    /// </summary>
    /// <exception cref="Exception">Installed hardware does not support high-resolution performance counters. <see cref="http://msdn.microsoft.com/en-us/library/ms644905(VS.85).aspx"/></exception>
    public class TickTimer
    {
        private readonly long _startTime;
        private static readonly double Resolution;
        private long _stopTime;
        public const long TicksPerMillisecond = 10000;
        private static readonly long SessionStartTime;
        private static readonly long SessionDateTimeUtcStartTicks;

        public TickTimer()
        {
            QueryPerformanceCounter(out _startTime);
        }

        static TickTimer()
        {
            long frequency;
            SessionDateTimeUtcStartTicks = DateTime.UtcNow.Ticks;
            QueryPerformanceCounter(out SessionStartTime);
            if (QueryPerformanceFrequency(out frequency))
            {
                Resolution = frequency / 1000.0d;
            }
            else
            {
                throw new Exception("Installed hardware does not support high-resolution performance counters. See http://msdn.microsoft.com/en-us/library/ms644905(VS.85).aspx");
            }
        }


        /// <summary>
        /// Stops measuring elapsed time.
        /// </summary>
        public void Stop()
        {
            QueryPerformanceCounter(out _stopTime);
        }




        /// <summary>
        /// Gets the elapsed ticks.
        /// </summary>
        /// <value>The elapsed ticks.</value>
        public long HighResElapsedTicks
        {
            get { return (_stopTime - _startTime); }
        }
        /// <summary>
        /// Gets the elapsed ticks.
        /// </summary>
        /// <value>The elapsed ticks.</value>
        /// <seealso cref="DateTime.Ticks"/>
        public long DateTimeElapsedTicks
        {
            get { return (long)((HighResElapsedTicks * TicksPerMillisecond) / Resolution); }
        }


        /// <summary>
        /// Gets the elapsed milliseconds.
        /// </summary>
        /// <value>The elapsed milliseconds.</value>
        public double ElapsedMilliseconds
        {
            get { return (HighResElapsedTicks / Resolution); }
        }

        /// <summary>
        /// Gets the <see cref="DateTime"/>, in Ticks, when the <see cref="TickTimer"/> started.
        /// </summary>
        /// <seealso cref="DateTime.Ticks"/>
        public long DateTimeUtcStartTicks
        {
            get
            {
                var ticksSinceSessionStart = (long)(((_startTime - SessionStartTime) * TicksPerMillisecond) / Resolution);
                return SessionDateTimeUtcStartTicks + ticksSinceSessionStart;
            }
        }


        [DllImport("kernel32.dll")]
        private static extern bool QueryPerformanceFrequency(out long value);


        [DllImport("kernel32.dll")]
        private static extern bool QueryPerformanceCounter(out long value);



    }
}
