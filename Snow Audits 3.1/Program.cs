using System.ComponentModel;
using System.Configuration;
using System.Data.SqlClient;
using System.Drawing.Drawing2D;
using System.Reflection.PortableExecutable;
using MenuBuilder;

namespace SnowAudit
{
    internal class Program
    {
        static void Main(string[] args)
        {
            UserInterface.DisplayInfo();
        ACTIONS:
            int actions = UserInterface.SelectAction();
        AUDITTYPE:
            int auditType;
            if (actions == 1)
            {
                auditType = UserInterface.AuditSelectAuditType();
            }
            else
            {
                UserInterface.NotImplemented();
                goto ACTIONS;
            }
            if (auditType == 2)
            {
                goto ACTIONS;
            }
            AuditProperties.SetAuditTypeInfo(auditType);
            var instanceType = UserInterface.SelectInstanceType();
            if (instanceType == 4)
            {
                goto AUDITTYPE;
            }
            AuditProperties.SetServerGroupInfo(AuditProperties.serverGroup);
            FileOperations.Initialize();
            DatabaseOperations.Initialize();
            UserInterface.ClearConsole();
            DatabaseOperations.ImportDataToAuditDB();
            UserInterface.ClearConsole();
            DatabaseOperations.PerformAudit();
            UserInterface.ClearConsole();
            UserInterface.WriteToConsole(@$"Complete, your results are at '{AuditProperties.outputFilePath}{AuditProperties.auditType} - {AuditProperties.serverGroup.ToUpper()} RESULTS.xlsx'.");
        }
    }
}