using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading;

namespace ExcelHelper
{
    /// <summary>
    /// TimeOutClass
    /// </summary>
    public class TimeOutClass
    {
        /// <summary>
        /// 超时函数
        /// </summary>
        /// <param name="action"></param>
        /// <param name="timeoutMilliseconds"></param>
        /// <returns></returns>
        public static bool CallWithTimeout(Action action, int timeoutMilliseconds)
        {
            Thread threadToKill = null;
            Action wrappedAction = () =>
            {
                threadToKill = Thread.CurrentThread;
                action();
            };
            IAsyncResult result = wrappedAction.BeginInvoke(null, null);
            if (result.AsyncWaitHandle.WaitOne(timeoutMilliseconds))
            {
                wrappedAction.EndInvoke(result);
                return false;
            }
            else
            {
                threadToKill.Abort();
                //throw new TimeoutException();
                return true;
            }
        }
        private readonly int timeOutSeconds = 5000;
        public long lastTicks;
        public long elaspedTicks;
        public TimeOutClass()
        {
            lastTicks = DateTime.Now.Ticks;
        }
        public TimeOutClass(int timeOutSeconds)
        {
            this.timeOutSeconds = timeOutSeconds;
            lastTicks = DateTime.Now.Ticks;
        }

        public bool isTimeOut()
        {
            elaspedTicks = DateTime.Now.Ticks - lastTicks;
            TimeSpan span = new TimeSpan(elaspedTicks);
            double diff = span.TotalSeconds;
            if (diff > timeOutSeconds)
                return true;
            else
                return false;
        }
    }
}
