using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.ServiceProcess;
using System.Text;
using System.Threading.Tasks;
using System.Media;
using System.Runtime.InteropServices;
using System.Threading;

namespace WindowsServiceSound
{
    public partial class Service1 : ServiceBase
    {
        public Service1()
        {

            InitializeComponent();
        }
        void Allthings()
        {
            var apps = new List<string>();
            Process[] Proc = Process.GetProcesses();
            foreach (Process p in Proc)
            {
                apps.Add(p.ProcessName);
            }
            foreach (string i in apps)
            {
                if (i == "wmplayer")
                {
                    int Level = 50;
                    IMMDeviceEnumerator deviceEnumerator = MMDeviceEnumeratorFactory.CreateInstance();
                    IMMDevice speakers;
                    const int eRender = 0;
                    const int eMultimedia = 1;
                    deviceEnumerator.GetDefaultAudioEndpoint(eRender, eMultimedia, out speakers);
                    object aepv_obj;
                    speakers.Activate(typeof(IAudioEndpointVolume).GUID, 0, IntPtr.Zero, out aepv_obj);
                    IAudioEndpointVolume aepv = (IAudioEndpointVolume)aepv_obj;
                    Guid ZeroGuid = new Guid();
                    int res = aepv.SetMasterVolumeLevelScalar(Level / 100f, ZeroGuid);

                }

                else { Time(); }
            }
        }
        public static void Time()
        {
            Service1 ex = new Service1();
            ex.StartTimer(100);

        }
        public void StartTimer(int dueTime)
        {
            Timer t = new Timer(new TimerCallback(myCallBack));
            t.Change(dueTime, 0);
        }
        private void myCallBack(object state)
        {
            Timer t = (Timer)state;
            t.Dispose();
            Allthings();
        }
       

        protected override void OnStart(string[] args)
        {
            Allthings();
        }

        protected override void OnStop()
        {

        }

        [ComImport]
        [Guid("A95664D2-9614-4F35-A746-DE8DB63617E6"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDeviceEnumerator
        {
            void _VtblGap1_1();
            int GetDefaultAudioEndpoint(int dataFlow, int role, out IMMDevice ppDevice);
        }
        private static class MMDeviceEnumeratorFactory
        {
            public static IMMDeviceEnumerator CreateInstance()
            {
                return (IMMDeviceEnumerator)Activator.CreateInstance(Type.GetTypeFromCLSID(new Guid("BCDE0395-E52F-467C-8E3D-C4579291692E"))); // a MMDeviceEnumerator
            }
        }
        [Guid("D666063F-1587-4E43-81F1-B948E807363F"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        private interface IMMDevice
        {
            int Activate([MarshalAs(UnmanagedType.LPStruct)] Guid iid, int dwClsCtx, IntPtr pActivationParams, [MarshalAs(UnmanagedType.IUnknown)] out object ppInterface);
        }

        [Guid("5CDF2C82-841E-4546-9722-0CF74078229A"), InterfaceType(ComInterfaceType.InterfaceIsIUnknown)]
        public interface IAudioEndpointVolume
        {
           
            int RegisterControlChangeNotify(IntPtr pNotify);
            
            int UnregisterControlChangeNotify(IntPtr pNotify);
            
            int GetChannelCount(ref uint pnChannelCount);
            
            int SetMasterVolumeLevel(float fLevelDB, Guid pguidEventContext);
            
            int SetMasterVolumeLevelScalar(float fLevel, Guid pguidEventContext);
            
            int GetMasterVolumeLevel(ref float pfLevelDB);

            int GetMasterVolumeLevelScalar(ref float pfLevel);
        }
    }


}

