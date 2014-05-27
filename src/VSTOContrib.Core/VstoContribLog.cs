using System;
using System.Collections.Generic;
using VSTOContrib.Core.Annotations;

namespace VSTOContrib.Core
{
    public static class VstoContribLog
    {
        static readonly List<Action<string>> Loggers = new List<Action<string>>();
        static VstoContribLogLevel level = VstoContribLogLevel.Info;

        public static void ToDebug()
        {
            Loggers.Add(l => System.Diagnostics.Debug.WriteLine(l));
        }

        public static void ToTrace()
        {
            Loggers.Add(l => System.Diagnostics.Trace.WriteLine(l));
        }

        public static void SetLevel(VstoContribLogLevel logLevel)
        {
            level = logLevel;
        }

        internal static void Debug([InstantHandle] Action<LogWriter> write)
        {
            if (level == VstoContribLogLevel.Debug)
                write((message, args) => Log("[Debug] ", message, args));
        }

        internal static void Info([InstantHandle] Action<LogWriter> write)
        {
            if (level != VstoContribLogLevel.Warn)
                write((message, args) => Log("[Info] ", message, args));
        }

        internal static void Warn([InstantHandle] Action<LogWriter> write)
        {
            write((message, args) => Log("[Warn] ", message, args));
        }

        static void Log(string level, string message, params object[] args)
        {
            foreach (var logger in Loggers)
            {
                logger(string.Concat(level, string.Format(message, args)));
            }
        }
    }
}