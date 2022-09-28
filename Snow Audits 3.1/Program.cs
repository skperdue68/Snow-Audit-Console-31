using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.Diagnostics.PerformanceData;
using System.Drawing.Drawing2D;
using System.Reflection.PortableExecutable;
using MenuBuilder;

namespace SnowAudit
{
    internal class Program
    {
        static void Main(string[] args)
        {
            //Display program information
            UserInterface.DisplayInfo();

        // Get the type of action to perform (perform audit, maintain exemptions etc).  Currently only Perform Audit is impemented.
        GETACTION:
            int action = UserInterface.SelectAction();
            if (action == 2)
            {
                UserInterface.NotImplemented();
                goto GETACTION;
            }

        //Get the type of Audit you are performing currently only System Properties is created, others may be added in a future release.
        GETAUDITTYPE:
            int auditType;
            auditType = UserInterface.SelectAuditType();
            if (auditType == 2)
            {
                goto GETACTION;
            }
            AuditProperties.SetAuditTypeParameters(auditType);

        //Get the server group for which you are performing the audit.
        GETSERVERGROUP:
            var serverGroup = UserInterface.SelectServerGroup();
            if (serverGroup == 4)
            {
                goto GETAUDITTYPE;
            }
            AuditProperties.SetServerGroupParameters(AuditProperties.serverGroup);

            UserInterface.ClearConsole();

            //Check for necessary directories and files based on instance type selected above.  Creates directories as necessary.
            FileOperations.Initialize();

            //Create Databases as necessary and clear/create tables as needed.
            DatabaseOperations.Initialize();

            UserInterface.AuditReview();

            DataSet ds = new DataSet();
            // Read the Excel File
            ds = DatabaseOperations.ReadExcel(ds);

            //Save Data to Database
            DatabaseOperations.AutoSqlBulkCopy(ds);

            //Run SQL aganst database to create the Audit Results and save outut
            UserInterface.ClearConsole();
            DatabaseOperations.PerformAudit();
            UserInterface.Logger("");
            UserInterface.Logger(@$"Audit Completed.  Results available as '{AuditProperties.outputFilePath}{AuditProperties.auditType} - {AuditProperties.serverGroup.ToUpper()} RESULTS.xlsx'.");
            UserInterface.Pause();
            goto GETSERVERGROUP;
        }
    }
}