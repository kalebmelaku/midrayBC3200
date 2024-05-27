namespace Mindray_BC3200
{
    using System;
    using System.ServiceProcess;

    internal static class Program
    {
        private static void Main()
        {
            ServiceBase[] services = new ServiceBase[] { new Service1() };
            ServiceBase.Run(services);
        }
    }
}

