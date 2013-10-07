using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Runtime.InteropServices;
using EasyHook;

namespace FileMonInject
{
    
    [StructLayout(LayoutKind.Sequential, CharSet = CharSet.Unicode)]
    public class MULTI_QI
    {
        //[MarshalAs(UnmanagedType.LPStruct)]
        public Guid IID;
        [MarshalAs(UnmanagedType.Interface)]
        public IntPtr pItf;
        public uint hr;
    }


    public partial class Main : EasyHook.IEntryPoint
    {
        //HRESULT CoCreateInstanceEx(
        //  _In_     REFCLSID rclsid,
        //  _In_     IUnknown *punkOuter,
        //  _In_     DWORD dwClsCtx,
        //  _In_     COSERVERINFO *pServerInfo,
        //  _In_     DWORD dwCount,
        //  _Inout_  MULTI_QI *pResults
        //);

        //DWORD WINAPI GetWindowThreadProcessId(
        //  _In_       HWND hWnd,
        //  _Out_opt_  LPDWORD lpdwProcessId
        //);

        [DllImport("user32.dll", CharSet = CharSet.Unicode, PreserveSig = true)]
        public static extern uint GetWindowThreadProcessId(
            [In] int hWnd,
            [Out] out uint lpdwProcessId);

        const int E_NOINTERFACE = unchecked((int)0x80004002);
        const int S_OK = 0x00000000;

        [DllImport("ole32.dll", CharSet = CharSet.Unicode, PreserveSig = true, SetLastError = true)]
        public static extern uint CoCreateInstanceEx(
            [In, MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
            [In] IntPtr punkOuter,
            [In] RegistrationClassContext dwClsCtx,
            [In] ref IntPtr pServerInfo,
            [In] uint dwCount,
            //[In, Out] IntPtr pResults
            [In, Out] IntPtr pResults
        );

        [UnmanagedFunctionPointer(CallingConvention.StdCall, CharSet=CharSet.Unicode,SetLastError=true)]
        public delegate uint DCoCreateInstanceEx(
           [In, MarshalAs(UnmanagedType.LPStruct)] Guid rclsid,
           [In] IntPtr punkOuter,
           [In] RegistrationClassContext dwClsCtx,
           [In] ref IntPtr pServerInfo,
           [In] uint dwCount,
           [In, Out] IntPtr pResults
       );

        public static uint CoCreateInstanceEx_Hook(
                Guid rclsid,
                IntPtr punkOuter,
                RegistrationClassContext dwClsCtx,
                ref IntPtr pServerInfo,
                uint dwCount,
                IntPtr pResults
            )
        {

            uint result = CoCreateInstanceEx(
                rclsid,
                punkOuter,
                dwClsCtx,
                ref pServerInfo,
                dwCount,
                pResults
            );

            if (dwCount > 0)
            {
                //To access the MULTI_QI array from inside the EasyHook delegate,
                //we have to manually marshal it because it was causing crashes.
                MULTI_QI[] pResultsObj = new MULTI_QI[dwCount];

                for (int i = 0; i < dwCount; i++)
                {
                    MULTI_QI qi = new MULTI_QI();
                    IntPtr ptrToIID = Marshal.ReadIntPtr(pResults);
                    qi.IID = (Guid)Marshal.PtrToStructure(ptrToIID, typeof(Guid));
                    pResults = IntPtr.Add(pResults, IntPtr.Size);
                    qi.pItf = Marshal.ReadIntPtr(pResults); //A pointer to the interface requested in pIID. This member must be NULL on input.
                    
                    pResults = IntPtr.Add(pResults, IntPtr.Size);

                    switch (IntPtr.Size)
                    {
                        case 8: //64-bit:
                            qi.hr = (uint)Marshal.ReadInt64(pResults);
                            pResults = IntPtr.Add(pResults, sizeof(UInt64)); 
                            break;
                        case 4: //32-bit:
                             qi.hr = (uint)Marshal.ReadInt32(pResults);
                            pResults = IntPtr.Add(pResults, sizeof(UInt32));
                            break;
                        default:
                            break; //unknown platform
                    }


                    //68cce6c0... is Access._Application interface
                    if (qi.hr == S_OK && qi.IID == Guid.Parse("68cce6c0-6129-101b-af4e-00aa003f0f07"))
                    {
                        Microsoft.Office.Interop.Access.Application app;
                        
                        app = Marshal.GetObjectForIUnknown(qi.pItf) as Microsoft.Office.Interop.Access.Application;

                        if (app != null)
                        {
                            uint ProcID = 0;
                            GetWindowThreadProcessId(app.hWndAccessApp(), out ProcID);

                            try
                            {
                                lock (This.AccessInstances)
                                {
                                    if (ProcID != 0)
                                    {
                                        This.AccessInstances.Enqueue(ProcID);
                                    }
                                }
                            }
                            catch { }
                        }
                    }

                    pResultsObj[i] = qi;
                }
            }

            return result;
        }
    }
}
