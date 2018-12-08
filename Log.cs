using System;
using System.Threading;

namespace ClassLibrary
{
    public class Log
    {
        public static void Write(String port, String msg)
        {
            DateTime now = DateTime.Now;
            Int32 threadId = Thread.CurrentThread.ManagedThreadId;
            String msg1 = String.Format("{0} {1:yyyy.MM.dd HH:mm:ss} {2}> {3}", port, now, threadId, msg);
            Console.WriteLine(msg1);
            SqlServer.LogWrite(msg, port);
        }
    }
}
