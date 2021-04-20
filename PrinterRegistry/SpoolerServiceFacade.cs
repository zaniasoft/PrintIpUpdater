using System;
using System.IO;
using System.ServiceProcess;

// <copyright file="SpoolerService.cs" company="">
// Copyright (c) 2019 All Rights Reserved
// </copyright>
// <author>Apichart Fuengaksorn</author>
// <date>11/11/2019 11:00:00 PM </date>
// <summary>Class to restart Spooler service</summary>

namespace PrintIpUpdater
{
    class SpoolerServiceFacade
    {
        readonly static ServiceController service = new ServiceController("Spooler");
        public static void Stop()
        {
            // Stop the spooler.
            if ((!service.Status.Equals(ServiceControllerStatus.Stopped)) &&
                (!service.Status.Equals(ServiceControllerStatus.StopPending)))
            {
                service.Stop();
                service.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(8));
            }
        }

        public static void Start()
        {
            // Start the spooler.
            service.Start();
            service.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(8));
        }

        public static void ClearSpoolCache()
        {
            System.IO.DirectoryInfo di = new DirectoryInfo("C:\\Windows\\System32\\spool\\PRINTERS");

            foreach (FileInfo file in di.GetFiles())
            {
                file.Delete();
            }
            foreach (DirectoryInfo dir in di.GetDirectories())
            {
                dir.Delete(true);
            }
        }

        public static void Restart()
        {
            Stop();
            ClearSpoolCache();
            Start();
        }

    }
}
