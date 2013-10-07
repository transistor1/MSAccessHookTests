using System;
using System.Collections.Generic;
using System.Text;
using System.Threading;
using System.Runtime.InteropServices;
using EasyHook;
using System.Collections.Concurrent;

namespace FileMonInject
{
    public partial class Main : EasyHook.IEntryPoint
    {
        FileMon.FileMonInterface Interface;
        LocalHook CreateFileHook;
        LocalHook TestHook;
        ConcurrentQueue<String> FileQueue = new ConcurrentQueue<String>();
        static ConcurrentQueue<uint> AccessInstances = new ConcurrentQueue<uint>();

        public Main(
            RemoteHooking.IContext InContext,
            String InChannelName)
        {
            // connect to host...
            Interface = RemoteHooking.IpcConnectClient<FileMon.FileMonInterface>(InChannelName);

            Interface.Ping();
        }

        public void Run(
            RemoteHooking.IContext InContext,
            String InChannelName)
        {
            // install hook...
            try
            {

                CreateFileHook = LocalHook.Create(
                    LocalHook.GetProcAddress("kernel32.dll", "CreateFileW"),
                    new DCreateFile(CreateFile_Hooked),
                    this);

                CreateFileHook.ThreadACL.SetExclusiveACL(new Int32[] { 0 });

                TestHook = LocalHook.Create(
                    LocalHook.GetProcAddress("ole32.dll", "CoCreateInstanceEx"),
                    new DCoCreateInstanceEx(CoCreateInstanceEx_Hook),
                    this);

                TestHook.ThreadACL.SetExclusiveACL(new Int32[] { 0 });
            }
            catch (Exception ExtInfo)
            {
                Interface.ReportException(ExtInfo);

                return;
            }

            Interface.IsInstalled(RemoteHooking.GetCurrentProcessId());

            RemoteHooking.WakeUpProcess();

            // wait for host process termination...
            try
            {
                while (true)
                {
                    Thread.Sleep(500);

                    if (AccessInstances.Count > 0)
                    {
                        uint[] Instances;
                        lock (AccessInstances)
                        {
                            Instances = AccessInstances.ToArray();

                            //FileQueue.Clear();
                            uint tmp;
                            while (AccessInstances.Count > 0)
                                AccessInstances.TryDequeue(out tmp);

                           
                        }

                        Interface.OnSpawnAccess(RemoteHooking.GetCurrentProcessId(), Instances);            
                    }

                    // transmit newly monitored file accesses...
                    if (FileQueue.Count > 0)
                    {
                        String[] Package = null;

                        lock (FileQueue)
                        {
                            Package = FileQueue.ToArray();

                            //FileQueue.Clear();
                            string tmp;
                            while (FileQueue.Count > 0)
                                FileQueue.TryDequeue(out tmp);
                        }

                        Interface.OnCreateFile(RemoteHooking.GetCurrentProcessId(), Package);
                    }
                    else
                        Interface.Ping();
                }
            }
            catch
            {
                // Ping() will raise an exception if host is unreachable
            }
        }

        [UnmanagedFunctionPointer(CallingConvention.StdCall,
            CharSet = CharSet.Unicode,
            SetLastError = true)]
        delegate IntPtr DCreateFile(
            String InFileName,
            UInt32 InDesiredAccess,
            UInt32 InShareMode,
            IntPtr InSecurityAttributes,
            UInt32 InCreationDisposition,
            UInt32 InFlagsAndAttributes,
            IntPtr InTemplateFile);

        // just use a P-Invoke implementation to get native API access from C# (this step is not necessary for C++.NET)
        [DllImport("kernel32.dll",
            CharSet = CharSet.Unicode,
            SetLastError = true,
            CallingConvention = CallingConvention.StdCall)]
        static extern IntPtr CreateFile(
            String InFileName,
            UInt32 InDesiredAccess,
            UInt32 InShareMode,
            IntPtr InSecurityAttributes,
            UInt32 InCreationDisposition,
            UInt32 InFlagsAndAttributes,
            IntPtr InTemplateFile);

        // this is where we are intercepting all file accesses!
        static IntPtr CreateFile_Hooked(
            String InFileName,
            UInt32 InDesiredAccess,
            UInt32 InShareMode,
            IntPtr InSecurityAttributes,
            UInt32 InCreationDisposition,
            UInt32 InFlagsAndAttributes,
            IntPtr InTemplateFile)
        { 
            
            try
            {
                Main This = (Main)HookRuntimeInfo.Callback;

                lock (This.FileQueue)
                {
                    This.FileQueue.Enqueue("[" + RemoteHooking.GetCurrentProcessId() + ":" + 
                        RemoteHooking.GetCurrentThreadId() +  "]: \"" + InFileName + "\"");
                }
            }
            catch
            {
            }

            // call original API...
            return CreateFile(
                InFileName,
                InDesiredAccess,
                InShareMode,
                InSecurityAttributes,
                InCreationDisposition,
                InFlagsAndAttributes,
                InTemplateFile);
        }
    }
}
