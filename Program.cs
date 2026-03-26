using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;

namespace CheckPort
{
    internal class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("[*] Looking for available ports..");
            string port = checkPorts(new string[] { "SYSTEM", "ANY" }).ToString();
            if (port == "-1")
            {
                Console.WriteLine("[-] No available ports found");
                Console.WriteLine("[-] Firewall will block our COM connection. Exiting");
                return;
            }
            Console.WriteLine("[*] Using port: " + port);
        }

        public static int checkPorts(string[] names)
        {
            IPGlobalProperties ipGlobalProperties = IPGlobalProperties.GetIPGlobalProperties();
            IPEndPoint[] tcpConnInfoArray = ipGlobalProperties.GetActiveTcpListeners();
            List<int> tcpPorts = tcpConnInfoArray.Select(i => i.Port).ToList();
            foreach (string name in names)
            {
                for (int i = 10; i < 65535; i++)
                {
                    if (checkPort(i, name) && !tcpPorts.Contains(i))
                    {
                        Console.WriteLine("[*] {0} Is allowed through port {1}", name, i);
                        return i;
                    }
                }
            }
            return -1;
        }

        public static bool checkPort(int port, string name = "SYSTEM")
        {
            try
            {
                // NET_FW_IP_VERSION_ANY = 2
                // NET_FW_IP_PROTOCOL_TCP = 6
                Type fwMgrType = Type.GetTypeFromProgID("HNetCfg.FwMgr");
                object mgr = Activator.CreateInstance(fwMgrType);

                // Check if firewall is enabled
                object localPolicy = fwMgrType.InvokeMember("LocalPolicy",
                    System.Reflection.BindingFlags.GetProperty, null, mgr, null);
                object currentProfile = localPolicy.GetType().InvokeMember("CurrentProfile",
                    System.Reflection.BindingFlags.GetProperty, null, localPolicy, null);
                bool firewallEnabled = (bool)currentProfile.GetType().InvokeMember("FirewallEnabled",
                    System.Reflection.BindingFlags.GetProperty, null, currentProfile, null);

                if (!firewallEnabled)
                    return true;

                // IsPortAllowed(name, ipVersion, portNumber, localAddress, ipProtocol, allowed, restricted)
                object[] parameters = new object[] { name, 2, port, "", 6, null, null };
                fwMgrType.InvokeMember("IsPortAllowed",
                    System.Reflection.BindingFlags.InvokeMethod, null, mgr, parameters);

                return (bool)parameters[5];
            }
            catch
            {
                return false;
            }
        }
    }
}
