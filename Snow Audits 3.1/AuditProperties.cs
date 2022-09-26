using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Security.Policy;
using System.Text;
using System.Threading.Tasks;

namespace SnowAudit
{
    internal class AuditProperties
    {
        public static string auditType = String.Empty;
        public static string serverGroup = String.Empty;
        public static string productionServer = String.Empty;
        public static List<string> servers = new List<string>();
        public static string dbAuditPrefix = String.Empty;
        public static string dbServer = String.Empty;
        public static string inputFilePath = @"C:\Audits\";
        public static string outputFilePath = @"C:\Audits\Output\";
        public static string exemptionsDatabase = "auditExemptions";
        public static string exemptionsTable = "exemptions";
        public static string dbAuditTableStructure = String.Empty;
        public static int outputFreezeRows = 0;
        public static int outputFreezeCols = 0;
        public static Dictionary<string, int> columnWidths = new Dictionary<string, int>();
        public static bool wrapRows = false;

        public static void SetAuditTypeInfo(int auditType)
        {
            switch (auditType)
            {
                case 1:
                    {
                        AuditProperties.dbAuditPrefix = "sys_prop";
                        AuditProperties.dbAuditTableStructure = "([Name] nvarchar(255) NULL,[Value] nvarchar(max) NULL,[Type] nvarchar(255) NULL,[Application] nvarchar(255) NULL,[Description] nvarchar(max) NULL,[Updated] datetime NULL,[Updated by] nvarchar(255) NULL)";
                        AuditProperties.outputFreezeCols = 2;
                        AuditProperties.outputFreezeRows = 1;
                        AuditProperties.columnWidths.Clear();
                        AuditProperties.columnWidths.Add("A", 20);
                        AuditProperties.columnWidths.Add("B", 50);
                        AuditProperties.columnWidths.Add("C", 20);
                        AuditProperties.columnWidths.Add("D", 20);
                        AuditProperties.columnWidths.Add("E", 50);
                        AuditProperties.columnWidths.Add("F", 50);
                        AuditProperties.columnWidths.Add("G", 50);
                        AuditProperties.columnWidths.Add("H", 50);
                        AuditProperties.columnWidths.Add("I", 50);
                        AuditProperties.columnWidths.Add("J", 50);
                        AuditProperties.wrapRows = true;
                        break;
                    }
            }
        }

        public static void SetServerGroupInfo(string serverGroup)
        {
            switch (serverGroup)
            {
                case "BDAS":
                    {
                        productionServer = "attbdas";
                        dbServer = "_bdas_";
                        servers.Clear();
                        servers.Add("attbdas");
                        servers.Add("attbdasbeta");
                        servers.Add("attbdastest");
                        servers.Add("attbdasupgradetest");
                        servers.Add("attbdasdev");
                        servers.Add("attbdasdev2");
                        servers.Add("attbdasdev3");
                        servers.Add("attbdasdev5");
                        servers.Add("attbdasdev6");
                        servers.Add("attbdasdev7");
                        break;
                    }
                case "FUSION":
                    {
                        productionServer = "attfusion";
                        dbServer = "_fusion_";
                        servers.Clear();
                        servers.Add("attfusion");
                        servers.Add("attfusiontest");
                        servers.Add("attfusiondev");
                        break;
                    }
                case "FEDGOV":
                    {
                        productionServer = "attfedgov1";
                        dbServer = "_fedgov_";
                        servers.Clear();
                        servers.Add("attfedgov1");
                        servers.Add("attfedgov1beta");
                        servers.Add("attfedgov1test");
                        servers.Add("attfedgov1dev");
                        servers.Add("attfedgov1dev2");
                        break;
                    }
            }
        }
    }
}
