using CommandLine;
using Microsoft.Win32;
using System;
using System.Collections.Generic;

// <copyright file="PrintIpUpdater.cs" company="">
// Copyright (c) 2019 All Rights Reserved
// </copyright>
// <author>Apichart Fuengaksorn</author>
// <date>11/11/2019 11:00:00 PM </date>
// <summary>Console app to update printer's IP address and restart spooler service</summary>
namespace PrintIpUpdater
{
    class PrintIpUpdater
    {
        const int SUCCESS = 0;
        const int ERROR_COMMAND_ARGS = 1;
        const int ERROR_READ_REGISTRY = 2;
        const int ERROR_WRITE_REGISTRY = 3;
        const int ERROR_RESTART_SPOOLER_SERVICE = 4;
        const int ERROR_STOP_SPOOLER_SERVICE = 5;
        const int ERROR_START_SPOOLER_SERVICE = 6;

        private static readonly string REGISTRY_KEY_BASE_PATH = "SYSTEM\\ControlSet001\\Control\\Print\\Monitors\\";

        /// <summary>
        /// This C# code reads a key from the windows registry.
        /// </summary>
        /// <param name="keyName">
        /// <returns></returns>
        public static string Read(string keyName)
        {
            string result = null;
            RegistryKey sk = Microsoft.Win32.Registry.LocalMachine.OpenSubKey(REGISTRY_KEY_BASE_PATH + options.Monitor + "\\Ports\\" + options.Port);
            if (sk != null)
            {
                try
                {
                    result = sk.GetValue(keyName.ToUpper()).ToString();
                }
                catch (NullReferenceException) { }
            }
            return result;
        }

        /// <summary>
        /// This C# code writes a key to the windows registry.
        /// </summary>
        /// <param name="keyName">
        /// <param name="value">
        public static void Write(string keyName, string value)
        {
            RegistryKey sk1 = Microsoft.Win32.Registry.LocalMachine.CreateSubKey(REGISTRY_KEY_BASE_PATH + options.Monitor + "\\Ports\\" + options.Port);
            sk1.SetValue(keyName.ToUpper(), value);
        }

        public class Options
        {
            [Option('m', "monitor", Default = "ZDesigner Port Monitor", HelpText = "Monitor Name (Ex. ZDesigner Port Monitor)")]
            public string Monitor { get; set; }

            [Option('p', "port", Default = "LAN_0", HelpText = "Port (Ex. LAN_0)")]
            public string Port { get; set; }

            [Option('i', "ipaddr", Required = true, HelpText = "IP Address")]
            public string Ipaddr { get; set; }

            [Option('d', "mode", Default = "Normal", HelpText = "Mode (Normal, UpdateIP, Restart)")]
            public string Mode { get; set; }

        }

        static Options options = new Options();

        static void Main(string[] args)
        {
            CommandLine.Parser.Default.ParseArguments<Options>(args)
              .WithParsed<Options>(opts => RunOptionsAndReturnExitCode(opts))
              .WithNotParsed<Options>((errs) => HandleParseError(errs));

            string currentIpAddrValue = Read("IPAddress");

            if (currentIpAddrValue == null)
            {
                Console.WriteLine("Error : Can not read current IP address");
                Environment.Exit(ERROR_READ_REGISTRY);
            }

            if (options.Mode.Equals("Restart"))
            {
                Console.Write("Only Restarting spooler..");
                try
                {
                    SpoolerServiceFacade.Restart();
                    Console.WriteLine("Done");
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Error : " + ex.Message + " (" + ex.InnerException + ")");
                    Environment.Exit(ERROR_RESTART_SPOOLER_SERVICE);
                }
                Environment.Exit(SUCCESS);
            }

            if (!currentIpAddrValue.Equals(options.Ipaddr))
            {
                List<String> ipaddr = new List<String>();
                if (options.Mode.Equals("Normal"))
                {
                    ipaddr.Add("127.0.0.1");
                }
                ipaddr.Add(options.Ipaddr);

                ipaddr.ForEach(delegate (String ip)
                {
                    // Stop Spooler service
                    if (options.Mode.Equals("Normal"))
                    {
                        Console.Write("Stopping spooler service..");
                        try
                        {
                            SpoolerServiceFacade.StopAndClearSpoolCache();
                            Console.WriteLine("Done");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error : " + ex.Message + " (" + ex.InnerException + ")");
                            Environment.Exit(ERROR_STOP_SPOOLER_SERVICE);
                        }
                    }

                    // Change IP address
                    Console.Write("Changing IP address to " + ip + "..");
                    try
                    {
                        Write("IPAddress", ip);
                    }
                    catch (UnauthorizedAccessException ex)
                    {
                        Console.WriteLine("Failed");
                        Console.WriteLine("Error : " + ex.Message);
                        Console.WriteLine("Try to run as administrator");
                        Environment.Exit(ERROR_WRITE_REGISTRY);
                    }

                    // Verify IP address
                    string ipAddrValue = Read("IPAddress");
                    if (String.Equals(ipAddrValue, ip))
                    {
                        Console.WriteLine("Done");
                    }
                    else
                    {
                        Console.WriteLine("Error : Failed to write ip address");
                        Environment.Exit(ERROR_WRITE_REGISTRY);
                    }

                    // Start Spooler service
                    if (options.Mode.Equals("Normal"))
                    {
                        Console.Write("Starting spooler..");
                        try
                        {
                            SpoolerServiceFacade.Start();
                            Console.WriteLine("Done");
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine("Error : " + ex.Message + " (" + ex.InnerException + ")");
                            Environment.Exit(ERROR_START_SPOOLER_SERVICE);
                        }
                    }
                });
            }
            Environment.Exit(SUCCESS);
        }

        private static void HandleParseError(IEnumerable<Error> errs)
        {
            Environment.Exit(ERROR_COMMAND_ARGS);
        }

        private static void RunOptionsAndReturnExitCode(Options opts)
        {
            options = opts;
        }
    }
}
