using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Xml.Linq;
using Utility.Common;

namespace Utility.Common
{
    public  class PublicTool
    {
        const int PORT_REMOTE = 8888;
        const int PORT_LOCAL = 9999;
        public static UdpClient localUdp = new UdpClient(PORT_LOCAL);
        public static IPEndPoint GetRemoteEp()
        {
            try
            {
                Utility.Log.TextLog.WritwLog("GetRemoteEp");
                foreach (var sti in Dns.GetHostAddresses(Dns.GetHostName()))
                {
                    Utility.Log.TextLog.WritwLog("GetRemoteEp_"+sti.ToString());
                    if (System.Text.RegularExpressions.Regex.IsMatch(sti.ToString(), "192.168."))
                    {
                        return new IPEndPoint(sti, PORT_REMOTE);
                    }
                }
                return null;
            }
            catch (Exception e)
            {
                Utility.Log.TextLog.WritwLog(e.Message);
                return null;
            }
        }
    }
}
