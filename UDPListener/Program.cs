using System;
using System.Diagnostics;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;

namespace UDPListener
{
    class Program
    {
        private static readonly log4net.ILog filelog = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        static int Main(string[] args)
        {
            try
            {
                if (args.Length != 1)
                {
                    Console.WriteLine("One parameter required, Port Number.");
                    return 1;
                }

                if (!int.TryParse(args[0], out int PortNum))
                {
                    throw new ArgumentException("Unable to parse Port Number.");
                }

                string OutputFolder = Path.GetTempPath();
                if (!Directory.Exists(OutputFolder))
                {
                    throw new ArgumentException("Temp Destination Folder '" + OutputFolder + "' does not exist.");
                }

                log4net.GlobalContext.Properties["XLM_ProcessID"] = Process.GetCurrentProcess().Id.ToString();
                log4net.GlobalContext.Properties["XLM_OutputFolder"] = OutputFolder;

                string log4netConfigPath = Path.Combine(FileHelpers.GetAssemblyFolder(), "log4net.config");
                FileInfo fi = new FileInfo(log4netConfigPath);
                log4net.Config.XmlConfigurator.Configure(fi);

                UdpClient listener = null;
                try
                {
                    listener = new UdpClient(PortNum);
                    IPEndPoint EP = new IPEndPoint(IPAddress.Any, PortNum);
                    Console.WriteLine("Listening on " + PortNum);
                    Console.WriteLine("Writing Rolling Log to Temp folder " + OutputFolder);

                    while (true)
                    {
                        byte[] b = listener.Receive(ref EP);
                        string data = Encoding.ASCII.GetString(b);
                        Console.WriteLine("UDP from {0}: {1}", EP.ToString(), data);
                        filelog.Info(data);
                    }
                }
                catch (Exception e)
                {
                    Console.WriteLine(e.ToString());
                }
                finally
                {
                    listener?.Close();
                }

            }
            catch (Exception e)
            {
                Console.WriteLine(e.ToString());
            }
            return 0;
        }

    }
}
