using System;
using System.Data;
using System.IO;
using System.Net;
using System.Text;
using System.Threading;

namespace Nskd
{
    public class HttpServer
    {
        private HttpListener listener;

        public delegate void RequestDelegate(HttpListenerContext context);
        public event RequestDelegate OnIncomingRequest;

        // Запуск сервера
        public void Start(String uriPrefix)
        {
            listener = new HttpListener();
            listener.Prefixes.Add(uriPrefix);
            listener.Start();
            // В бесконечном цикле
            while (true)
            {
                // Принимаем новых клиентов. 
                WaitCallback waitCallback = new WaitCallback(AcceptCompleted);
                // После того, как клиент был принят, он передается в новый поток с использованием пула потоков.
                ThreadPool.QueueUserWorkItem(waitCallback, listener.GetContext());
            }
        }

        /// <summary>
        /// This method is the callback method associated with Socket.AcceptAsync  
        /// operations and is invoked when an accept operation is complete.
        /// </summary>
        private void AcceptCompleted(Object context)
        {
            HttpListenerContext c = context as HttpListenerContext;
            if (c != null)
            {
                if (OnIncomingRequest != null)
                {
                    try
                    {
                        //  Process request and write responce
                        OnIncomingRequest(c);
                    }
                    catch (Exception e) { Console.Write(e.ToString()); }
                }
                c.Response.Close();
            }
        }

        /// <summary>
        /// Из входящего ответа от сайта делает исходящий ответ для пользователя.
        /// Копирует заголовки, а если есть, то и тело ответа.
        /// </summary>
        /// <param name="context"></param>
        /// <param name="incomingResponse"></param>
        public void SendResponse(HttpListenerContext context, HttpWebResponse incomingResponse)
        {
            WebHeaderCollection headers = incomingResponse.Headers;
            headers.Remove("Content-Length"); // Этот заголовок нельзя передавать в комплекте.
            context.Response.Headers = headers;
            if ((incomingResponse.ContentLength > 0) &&
                (context.Response.OutputStream.CanWrite))
            {
                Stream receiveStream = incomingResponse.GetResponseStream();
                MemoryStream ms = new MemoryStream();
                receiveStream.CopyTo(ms);
                context.Response.ContentLength64 = ms.Length; // Восстанавливаем Content-Length
                ms.Position = 0;
                ms.CopyTo(context.Response.OutputStream);
            }
            context.Response.OutputStream.Close();
        }

        /// <summary>
        /// Отправляет страницу с описанием ошибки и закрывает поток ответа.
        /// </summary>
        public void SendErrorPage(HttpListenerContext context, HttpStatusCode status, String msg = null)
        {
            String html = "<html><head><meta charset=\"UTF-8\"></head><body><h1>";
            html += ((int)status).ToString() + " " + status.ToString();
            if (!String.IsNullOrWhiteSpace(msg)) html += "<br / >" + msg;
            html += "</h1></body></html>";
            byte[] buff = Encoding.UTF8.GetBytes(html);
            //try
            {
                if (context.Response.OutputStream.CanWrite)
                {
                    context.Response.OutputStream.Write(buff, 0, buff.Length);
                }
            }
            //catch (Exception) { }
            context.Response.OutputStream.Close();
        }
    }
}
