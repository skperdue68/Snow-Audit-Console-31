using System;
using System.Collections;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Text;
using System.Threading.Tasks;
using System.Web;

namespace SnowAudit
{
    internal static class FileOperations
    {
        public static void Initialize()
        {
            CreateAuditDirectory();
            CheckAuditFiles();
        }

        public static void CreateAuditDirectory()
        {
            // Create the Audit Directory if it does not exist.
            Directory.CreateDirectory(AuditProperties.inputFilePath);
            Directory.CreateDirectory(AuditProperties.outputFilePath);
        }

        public static void CheckAuditFiles()
        {
            string inputFilePath = AuditProperties.inputFilePath;
            string productionServer = AuditProperties.productionServer;
            List<string> servers = AuditProperties.servers;
            List<string> inputFiles = new List<string>();
            if (!File.Exists(inputFilePath + productionServer + ".xlsx"))
            {
                UserInterface.ShowInputFileError(1);
            }
            foreach (string server in servers)
            {
                if (server != AuditProperties.productionServer)
                {
                    if (File.Exists(inputFilePath + server + ".xlsx"))
                    {
                        inputFiles.Add(server);
                    }
                }
            }
            if (inputFiles.Count < 1)
            {
                UserInterface.ShowInputFileError(2);
            }
            // Save list of servers based on Input Files found."
            AuditProperties.servers.Clear();
            AuditProperties.servers.Add(productionServer);
            AuditProperties.servers.AddRange(inputFiles.ToArray());
        }
    }
}
