using DocumentFormat.OpenXml.Office.CustomUI;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.IO;

namespace SnowAudit
{
    internal static class FileOperations
    {
        static string inputFilePath = AuditProperties.inputFilePath;
        static string outputFilePath = AuditProperties.outputFilePath;
        static string productionServer = AuditProperties.productionServer;
        static List<string> servers = AuditProperties.servers;


        internal static void Initialize()
        {
            CreateAuditDirectory();
            GatherAuditInputFiles();
        }

        internal static void CreateAuditDirectory()
        {
            // Create the Input and Output Directories if necessary.
            Directory.CreateDirectory(inputFilePath);
            Directory.CreateDirectory(outputFilePath);
        }

        //Check which file(s) are present in the input directory based on your instance to audit and create a list of the results.
        internal static void GatherAuditInputFiles()
        {

            UserInterface.LoggerChangeColors("FILE CHECK", ConsoleColor.White, ConsoleColor.Blue);
            UserInterface.LoggerChangeColors("", ConsoleColor.DarkBlue, ConsoleColor.White);
            List<string> inputFiles = new List<string>();

            //Look specifically for production server as this is always required.
            if (!File.Exists(inputFilePath + productionServer + ".xlsx"))
            {
                UserInterface.Logger($"{inputFilePath}{productionServer}.xlsx...NOT found!");
                UserInterface.ShowError("File 1", "");
            }
            else
            {
                inputFiles.Add(productionServer);
                UserInterface.Logger($"{inputFilePath}{productionServer}.xlsx...found!");
            }
            //Examine remaining possible server(s) files.
            foreach (string server in servers)
            {
                if (server != productionServer)
                {
                    UserInterface.Logger($"{inputFilePath}{server}.xlsx...", false);
                    if (File.Exists(inputFilePath + server + ".xlsx"))
                    {
                        inputFiles.Add(server);
                        UserInterface.Logger($"found!");
                    }
                    else
                    {
                        UserInterface.Logger($"NOT found!");
                    }
                }
            }

            // If there are not at least 2 input files (one being production and the other being one or more of the remaining options, display error and exit.
            if (inputFiles.Count < 2)
            {
                UserInterface.ShowError("File 2", "");
            }

            UserInterface.Logger("", false, System.ConsoleColor.White);
            UserInterface.Pause();

            // Save list of servers from the serverGroup with input files.
            AuditProperties.servers.Clear();
            AuditProperties.servers.AddRange(inputFiles.ToArray());
        }
    }
}
