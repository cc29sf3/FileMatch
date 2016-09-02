using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Manual_Import.Helper
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
    /// <summary>
    /// 描述：异步任务处理
    /// 作者：员战强
    /// </summary>
    public class UdpServer
    {
        
        // 定义端口
        private const int listenPort = 9999;
        private const int remotePort = 8888;
        public UdpClient sendUdpClient;
        private UdpState udpReceiveState = new UdpState();

        public UdpServer()
        {
            // 实例化
            sendUdpClient = new UdpClient(listenPort);
            
        }
      

        public string SendMessage(string message)
        {
            var name = Dns.GetHostName();
            var ssfs = Dns.GetHostAddresses(name);
            IPAddress localAddress = Dns.GetHostAddresses(Dns.GetHostName())[2];
            byte[] mess=Encoding.Default.GetBytes(message);
            int length=Encoding.Default.GetByteCount(message);
            IPEndPoint sendTo = new IPEndPoint(localAddress, remotePort);
            sendUdpClient.Send(mess, length,sendTo);

            udpReceiveState.udpClient = sendUdpClient;
            udpReceiveState.ipEndPoint = sendTo;
            
            byte[] receiveByte= sendUdpClient.Receive(ref sendTo);
            string receiveString = Encoding.ASCII.GetString(receiveByte);
            
            if (receiveString == "Y")
            {
                byte[] backByte = sendUdpClient.Receive(ref sendTo);
                string backString = Encoding.UTF8.GetString(backByte);
                return backString;
            }
            //sendUdpClient.Close();M
            
            return receiveString;
            
            //sendUdpClient.BeginReceive(ReceiveCallback, udpReceiveState);
        }

        public void ReceiveCallback(IAsyncResult iar)
        {
            UdpState udpState = iar.AsyncState as UdpState;
            if (iar.IsCompleted)
            {
                Byte[] receiveBytes =udpState.udpClient.EndReceive(iar, ref udpReceiveState.ipEndPoint);
                string receiveString = Encoding.ASCII.GetString(receiveBytes);
                Console.WriteLine(receiveString);
            }
        }
    }
}
