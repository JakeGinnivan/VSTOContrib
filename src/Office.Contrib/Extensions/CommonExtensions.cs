using System;
using System.Text;
using System.Threading;
using System.Diagnostics;

namespace Office.Contrib.Extensions
{
    /// <summary>
    /// Common extensions that are not Office specific, but are used by the library
    /// </summary>
    public static class CommonExtensions
    {
        /// <summary>
        /// Gets all messages (including inner exceptions) from the exception
        /// </summary>
        /// <param name="ex">The ex.</param>
        /// <returns></returns>
        public static string ToMessageStack(this Exception ex)
        {
            var sb = new StringBuilder();
            var currentException = ex;

            sb.AppendLine(ex.Message);

            var numExceptions = 1;

            while (currentException.InnerException != null)
            {
                currentException = currentException.InnerException;

                sb.AppendLine(string.Format("{0}-> {1}", new string(' ', numExceptions++*2),
                                            currentException.Message));
            }

            return sb.ToString();
        }

        /// <summary>
        /// Gets the full stack trace, including inner exceptions up to a specified level
        /// </summary>
        /// <param name="ex">The ex.</param>
        /// <param name="count">The count.</param>
        /// <returns></returns>
        public static string ToFullStackTrace(this Exception ex, int count)
        {
            if (ex == null)
                throw new ArgumentException();
            var sb = new StringBuilder();
            sb.Append(ex.StackTrace);

            var innerReferences = 0;
            var inner = ex.InnerException;
            while (inner != null && innerReferences < count)
            {
                sb.Insert(0, inner.StackTrace);
                inner = inner.InnerException;
                innerReferences++;
            }
            return sb.ToString();
        }

        /// <summary>
        /// Gets the full stack trace, including inner exceptions.
        /// </summary>
        /// <param name="ex">The ex.</param>
        /// <returns></returns>
        public static string ToFullStackTrace(this Exception ex)
        {
            return ex.ToFullStackTrace(50);
        }

        /// <summary>
        /// Starts the process, and reads all output from the process
        /// </summary>
        /// <param name="processStartInfo">The process start info.</param>
        /// <param name="outputDataRecieved">The output data recieved.</param>
        /// <returns></returns>
        public static int StartProcess(this ProcessStartInfo processStartInfo, DataReceivedEventHandler outputDataRecieved)
        {
            processStartInfo.UseShellExecute = false;
            processStartInfo.RedirectStandardOutput = true;
            processStartInfo.RedirectStandardError = true;
            processStartInfo.CreateNoWindow = false;

            var process = new Process {StartInfo = processStartInfo};

            if (outputDataRecieved != null)
            {
                process.OutputDataReceived += outputDataRecieved;
                process.ErrorDataReceived += outputDataRecieved;
            }
            process.Start();
            process.BeginOutputReadLine();
            process.BeginErrorReadLine();
            process.WaitForExit();

            return process.ExitCode;
        }

        /// <summary>
        /// Calls the action with a timeout.
        /// </summary>
        /// <param name="action">The action.</param>
        /// <param name="timeoutMilliseconds">The timeout milliseconds.</param>
        /// <exception cref="TimeoutException">If timeout occurs</exception>
        public static void CallWithTimeout(this Action action, int timeoutMilliseconds)
        {
            Thread threadToKill = null;
            Action wrappedAction = () =>
            {
                threadToKill = Thread.CurrentThread;
                action();
            };

            var result = wrappedAction.BeginInvoke(null, null);
            if (result.AsyncWaitHandle.WaitOne(timeoutMilliseconds))
            {
                wrappedAction.EndInvoke(result);
            }
            else
            {
                threadToKill.Abort();
                throw new TimeoutException("Action timed out");
            }
        }
    }
}
