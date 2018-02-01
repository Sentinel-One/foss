using Common.Logging;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Topshelf;
using Topshelf.Common.Logging;
using Topshelf.Ninject;

namespace SentinelOneNotificationService
{
    class Program
    {
        static void Main(string[] args)
        {

            // This will ensure that future calls to Directory.GetCurrentDirectory()
            // returns the actual executable directory and not something like C:\Windows\System32 
            Directory.SetCurrentDirectory(AppDomain.CurrentDomain.BaseDirectory);

            // Specify the base name, display name and description for the service, as it is registered in the services control manager.
            // This information is visible through the Windows Service Monitor
            const string serviceName = "SentinelOneNotificationService";
            const string displayName = "SentinelOne Notification Service";
            const string description = "Custom email notifications for SentinelOne alerts";

            HostFactory.Run(x =>
            {
                x.UseCommonLogging();
                x.UseNinject(new IocModule());

                x.Service<Service.S1Notifier>(sc =>
                {

                    sc.ConstructUsingNinject();

                    // the start and stop methods for the service
                    sc.WhenStarted((s, hostControl) => s.Start(hostControl));
                    sc.WhenStopped((s, hostControl) => s.Stop(hostControl));

                    // optional pause/continue methods if used
                    sc.WhenPaused((s, hostControl) => s.Pause(hostControl));
                    sc.WhenContinued((s, hostControl) => s.Continue(hostControl));

                    // optional, when shutdown is supported
                    sc.WhenShutdown((s, hostControl) => s.Shutdown(hostControl));

                });

                //=> Service Identity
                x.RunAsLocalSystem();
                //x.RunAs("username", "password"); // predefined user
                //x.RunAsPrompt(); // when service is installed, the installer will prompt for a username and password
                //x.RunAsNetworkService(); // runs as the NETWORK_SERVICE built-in account
                //x.RunAsLocalSystem(); // run as the local system account
                //x.RunAsLocalService(); // run as the local service account

                //=> Service Instalation - These configuration options are used during the service instalation
                x.StartAutomatically(); // Start the service automatically
                //x.StartAutomaticallyDelayed(); // Automatic (Delayed) -- only available on .NET 4.0 or later
                //x.StartManually(); // Start the service manually
                //x.Disabled(); // install the service as disabled

                //=> Service Configuration
                x.EnablePauseAndContinue(); // Specifies that the service supports pause and continue.
                x.EnableShutdown(); //Specifies that the service supports the shutdown service command.

                //=> Service Dependencies
                //=> Service dependencies can be specified such that the service does not start until the dependent services are started.
                //x.DependsOn("SomeOtherService");
                //x.DependsOnMsmq(); // Microsoft Message Queueing
                //x.DependsOnMsSql(); // Microsoft SQL Server
                //x.DependsOnEventLog(); // Windows Event Log
                //x.DependsOnIis(); // Internet Information Server

                x.SetDescription(description);
                x.SetDisplayName(displayName);
                x.SetServiceName(serviceName);

            });


        }
    }
}
