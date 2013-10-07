using System;
using System.Collections.Generic;
using System.Runtime.Remoting;
using System.Text;
using System.IO;
using EasyHook;
using System.Windows.Forms;

namespace FileMon
{
    public class FileMonInterface : MarshalByRefObject
    {
        public void IsInstalled(Int32 InClientPID)
        {
            Console.WriteLine("MSAccessHookTests has been installed in target {0}.\r\n", InClientPID);
        }

        public void OnCreateFile(Int32 InClientPID, String[] InFileNames)
        {
            for (int i = 0; i < InFileNames.Length; i++)
            {
                Console.WriteLine(InFileNames[i]);
            }
        }

        public void OnSpawnAccess(Int32 InClientPID, uint[] AccessInstances)
        {
            for (int i = 0; i < AccessInstances.Length; i++)
            {
                RemoteHooking.Inject(
                    (int)AccessInstances[i],
                    "MSAccessHookTests.exe",
                    "MSAccessHookTests.exe",
                    Program.ChannelName);
            }
        }

        public void ReportException(Exception InInfo)
        {
            Console.WriteLine("The target process has reported an error:\r\n" + InInfo.ToString());
        }

        public void Ping()
        {
        }
    }

    class Program
    {
        public static String ChannelName { get; set; }

        static void Main(string[] args)
        {
            Int32 TargetPID = 0;

            if ((args.Length != 1) || !Int32.TryParse(args[0], out TargetPID))
            {
                Console.WriteLine();
                Console.WriteLine("Usage: MSAccessHookTests %PID%");
                Console.WriteLine();

                return;
            }

            try
            {
                string _ChannelName = null;

                RemoteHooking.IpcCreateServer<FileMonInterface>(ref _ChannelName, WellKnownObjectMode.SingleCall);

                ChannelName = _ChannelName;

                RemoteHooking.Inject(
                    TargetPID,
                    "MSAccessHookTests.exe",
                    "MSAccessHookTests.exe",
                    ChannelName);
                
                Console.ReadLine();
            }
            catch (Exception ExtInfo)
            {
                Console.WriteLine("There was an error while connecting to target:\r\n{0}", ExtInfo.ToString());
            }
        }
    }
}