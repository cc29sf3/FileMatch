using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Utility.Log;

namespace FileMatch.Helper
{
    /// <summary>
    ///  UDP状态类
    /// </summary>
    public class UdpState
    {
        public UdpClient udpClient = null;
        public IPEndPoint ipEndPoint = null;
        public const int BufferSize = 1024;
        public byte[] buffer = new byte[BufferSize];
        public int counter = 0;
    }
   
    public class UdpServer
    {
        
        // 定义端口
        private const int listenPort = 9999;
        private const int remotePort = 8888;
        public UdpClient sendUdpClient;
        private static UdpServer udpServer;
        IPEndPoint sendTo;


        private UdpServer()
        {
            sendUdpClient = new UdpClient(listenPort);
        }

        public static UdpServer GetInstance()
        {
            if (udpServer == null)
            {
                udpServer = new UdpServer();
            }
            return udpServer;
        }

        public string SendMessage(string message)
        {
            lock (this)
            {
                var name = Dns.GetHostName();
                var ssfs = Dns.GetHostAddresses(name);
                IPAddress localAddress = Dns.GetHostAddresses(Dns.GetHostName())[2];
                byte[] mess = Encoding.Default.GetBytes(message);
                int length = Encoding.Default.GetByteCount(message);
                sendTo = new IPEndPoint(localAddress, remotePort);
                sendUdpClient.Send(mess, length, sendTo);
                //sendUdpClient.BeginSend(mess, length, SendCallback, 1);

                byte[] receiveByte = sendUdpClient.Receive(ref sendTo);
                string receiveString = Encoding.ASCII.GetString(receiveByte);
              

                if (receiveString == "Y")
                {
                    byte[] backByte = sendUdpClient.Receive(ref sendTo);
                    string backString = Encoding.Default.GetString(backByte);
                    return backString;
                }
                //sendUdpClient.Close();
                return receiveString;
            }
            
        }

        //public void ReceiveCallback(IAsyncResult iar)
        //{
        //    UdpState udpState = iar.AsyncState as UdpState;
        //    if (iar.IsCompleted)
        //    {
        //       // Byte[] receiveBytes = udpState.udpClient.EndReceive(iar, ref udpReceiveState.ipEndPoint);
        //       // string receiveString = Encoding.ASCII.GetString(receiveBytes);
        //        Console.WriteLine(receiveString);
        //    }
        //}

        //public void SendCallback(IAsyncResult iar)
        //{
        //    sendUdpClient.EndSend(iar);
        //    byte[] receiveByte = sendUdpClient.BeginReceive(
        //    string receiveString = Encoding.ASCII.GetString(receiveByte);
        //}

    }

    public class AsyncUdpClient
    {
        private const int listenPort = 9999;
        private const int remotePort = 8888;
        // 定义节点
        private IPEndPoint localEP = null;
        private IPEndPoint remoteEP = null;
        // 定义UDP发送和接收
        private UdpClient udpReceive = null;
        private UdpClient udpSend = null;
        //private UdpState udpSendState = null;
        private UdpState udpReceiveState = null;
        private int counter = 0;
        // 异步状态同步
       // private ManualResetEvent sendDone = new ManualResetEvent(false);
        private ManualResetEvent receiveDone = new ManualResetEvent(false);

        private string FinalBackString;

        public AsyncUdpClient()
        {
            var name = Dns.GetHostName();
            var ssfs = Dns.GetHostAddresses(name);
            // 本机节点
            localEP = new IPEndPoint(Dns.GetHostAddresses(Dns.GetHostName())[1], listenPort);
            // 远程节点
            remoteEP = new IPEndPoint(Dns.GetHostAddresses(Dns.GetHostName())[1], remotePort);
            // 实例化
            udpReceive = new UdpClient(localEP);
            udpSend = new UdpClient(listenPort);

            // 分别实例化udpSendState、udpReceiveState
            //udpSendState = new UdpState();
            //udpSendState.ipEndPoint = remoteEP;
            //udpSendState.udpClient = udpSend;

            udpReceiveState = new UdpState();
            udpReceiveState.ipEndPoint = remoteEP;
            udpReceiveState.udpClient = udpReceive;
        }

        public string SendMsg(string message)
        {
            //udpSend.Connect(remoteEP);
            Byte[]  sendBytes = Encoding.Default.GetBytes(message);
            // 调用发送回调函数
            try
            {
                udpSend.Send(sendBytes, sendBytes.Length, remoteEP);
            }
            catch (Exception f)
            {
                string s = f.Message;
            }
            ReceiveMessages();
            receiveDone.Reset();
            bool bbb= receiveDone.WaitOne();
            
            return FinalBackString;
        }

        public void ReceiveMessages()
        {
            lock (this)
            {
                Console.WriteLine("ReceiveMessages:" + Thread.CurrentThread.ManagedThreadId);
                udpReceive.BeginReceive(new AsyncCallback(ReceiveCallback), udpReceiveState);
                Thread.Sleep(100);
            }
        }

        // 接收回调函数
        public void ReceiveCallback(IAsyncResult iar)
        {
            UdpState udpState = iar.AsyncState as UdpState;
            if (iar.IsCompleted)
            {
                Console.WriteLine("ReceiveCallback:" + Thread.CurrentThread.ManagedThreadId);
                Byte[] receiveBytes = udpState.udpClient.EndReceive(iar, ref udpReceiveState.ipEndPoint);
                string receiveString = Encoding.UTF8.GetString(receiveBytes);
                if (receiveString == "Y")
                {
                    udpReceive.BeginReceive(new AsyncCallback(FinalReceiveCallback), udpReceiveState);
                }
                else
                {
                    FinalBackString = "N";
                    receiveDone.Set();
                }
               
            }
        }

        public void FinalReceiveCallback(IAsyncResult iar)
        {
            UdpState udpState = iar.AsyncState as UdpState;
            if (iar.IsCompleted)
            {
                Console.WriteLine("FinalReceiveCallback:" + Thread.CurrentThread.ManagedThreadId);
                Byte[] receiveBytes = udpState.udpClient.EndReceive(iar, ref udpReceiveState.ipEndPoint);
                FinalBackString= Encoding.UTF8.GetString(receiveBytes);
                receiveDone.Set();
            }
        }
    }
}
