using ClosedXML.Excel;
using DocumentFormat.OpenXml.Presentation;
using System.Collections.Generic;
using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Data.SqlTypes;
using System.Drawing;
using System.Xml.Linq;

namespace SnowAudit
{
    internal static class DatabaseOperations
    {
        private static string DBConnectionString = ConfigurationManager.ConnectionStrings["database"].ConnectionString;

        public static void Initialize()
        {
            CheckDBConnection();
            CheckForExemptionsDB(AuditProperties.exemptionsDatabase, AuditProperties.exemptionsTable);
            string dbName = AuditProperties.dbAuditPrefix + AuditProperties.dbServer + "audit";
            DropAuditDB(dbName);
            CreateAuditDB(dbName);
            DBConnectionString = DBConnectionString + $"Initial Catalog={dbName};";
            CreateAuditTable(dbName);
        }

        private static void CheckDBConnection()
        {
            using (SqlConnection dbConn = new SqlConnection(DBConnectionString))
            {
                try
                {
                    if (dbConn.State == System.Data.ConnectionState.Closed)
                    {
                        dbConn.Open();
                    }
                }
                catch (Exception ex)
                {
                    UserInterface.ShowDatabaseError(1, ex.Message);
                    System.Environment.Exit(1);
                }
                dbConn.Close();
            }
        }

        private static void CheckForExemptionsDB(string exemptionsDatabase, string exemptionsTable)
        {
            using (SqlConnection dbConn = new SqlConnection(DBConnectionString))
            {
                try
                {
                    if (dbConn.State == System.Data.ConnectionState.Closed)
                    {
                        dbConn.Open();
                    }
                    SqlCommand cmd = new SqlCommand(@$"IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = '{exemptionsDatabase}') BEGIN CREATE DATABASE [{exemptionsDatabase}] END", connection: dbConn);
                    cmd.ExecuteNonQuery();
                    cmd = new SqlCommand($"ALTER DATABASE [{exemptionsDatabase}] SET AUTO_CLOSE OFF", connection: dbConn);
                    cmd.ExecuteNonQuery();
                    dbConn.ChangeDatabase(exemptionsDatabase);
                    cmd = new SqlCommand(@$"IF NOT EXISTS (SELECT * FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_NAME = N'{exemptionsTable}') BEGIN CREATE TABLE [dbo].[exemptions]([ID] [bigint] IDENTITY(1,1) NOT NULL,[Name] [varchar](255) NOT NULL,[Audit_Type] [varchar](50) NOT NULL, CONSTRAINT [PK_exemptions] PRIMARY KEY CLUSTERED ([ID] ASC)WITH (PAD_INDEX = OFF, STATISTICS_NORECOMPUTE = OFF, IGNORE_DUP_KEY = OFF, ALLOW_ROW_LOCKS = ON, ALLOW_PAGE_LOCKS = ON, OPTIMIZE_FOR_SEQUENTIAL_KEY = OFF) ON [PRIMARY]) ON [PRIMARY] END", connection: dbConn);
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    UserInterface.ShowDatabaseError(2, ex.Message);
                    System.Environment.Exit(1);
                }
                dbConn.Close();
            }
        }

        private static void DropAuditDB(string dbName)
        {
            using (SqlConnection dbConn = new SqlConnection(DBConnectionString))
            {
                try
                {
                    if (dbConn.State == System.Data.ConnectionState.Closed)
                    {
                        dbConn.Open();
                    }
                    dbConn.ChangeDatabase("master");
                    SqlCommand cmd = new SqlCommand($"DROP DATABASE IF EXISTS {dbName}", connection: dbConn);
                    cmd.ExecuteNonQuery();
                }
                catch (Exception ex)
                {
                    UserInterface.ShowDatabaseError(3, ex.Message);
                    System.Environment.Exit(1);
                }
                dbConn.Close();
            }
        }

        private static void CreateAuditDB(string dbName)
        {
            using (SqlConnection dbConn = new SqlConnection(DBConnectionString))
            {
                try
                {
                    if (dbConn.State == System.Data.ConnectionState.Closed)
                    {
                        dbConn.Open();
                    }
                    SqlCommand cmd = new SqlCommand(@$"IF NOT EXISTS (SELECT * FROM sys.databases WHERE name = '{dbName}') BEGIN CREATE DATABASE [{dbName}] END", connection: dbConn);
                    cmd.ExecuteNonQuery();
                    cmd = new SqlCommand($"ALTER DATABASE [{dbName}] SET AUTO_CLOSE OFF", connection: dbConn);
                    cmd.ExecuteNonQuery();
                    dbConn.ChangeDatabase(dbName);
                }
                catch (Exception ex)
                {
                    UserInterface.ShowDatabaseError(4, ex.Message);
                    System.Environment.Exit(1);
                }
                dbConn.Close();
            }
        }

        private static void CreateAuditTable(string dbName)
        {
            using (SqlConnection dbConn = new SqlConnection(DBConnectionString))
            {
                try
                {
                    if (dbConn.State == System.Data.ConnectionState.Closed)
                    {
                        dbConn.Open();
                    }
                    dbConn.ChangeDatabase(dbName);
                    foreach (string server in AuditProperties.servers)
                    {
                        SqlCommand cmd = new SqlCommand($"CREATE TABLE {server} {AuditProperties.dbAuditTableStructure}", connection: dbConn);
                        cmd.ExecuteNonQuery();
                    }
                }
                catch (Exception ex)
                {
                    UserInterface.ShowDatabaseError(5, ex.Message);
                    System.Environment.Exit(1);
                }
                dbConn.Close();
            }
        }

        internal static void ImportDataToAuditDB()
        {
            try
            {
                string dbName = AuditProperties.dbAuditPrefix + AuditProperties.dbServer;
                foreach (string server in AuditProperties.servers)
                {
                    UserInterface.WriteToConsole($"Importing {server}.xlsx to Database...");
                    String excelConnString = String.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0\"", AuditProperties.inputFilePath + server + ".xlsx");
                    //Create Connection to Excel work book 
                    using (OleDbConnection excelConnection = new OleDbConnection(excelConnString))
                    {
                        //Create OleDbCommand to fetch data from Excel 
                        using (OleDbCommand cmd = new OleDbCommand("Select * FROM ['Page 1$']", excelConnection))
                        {
                            excelConnection.Open();
                            using (
                                OleDbDataReader dReader = cmd.ExecuteReader())
                            {
                                using (SqlBulkCopy sqlBulk = new SqlBulkCopy(DBConnectionString))
                                {
                                    sqlBulk.DestinationTableName = $"[{dbName}audit].[dbo].[{server}]";
                                    sqlBulk.WriteToServer(dReader);
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                UserInterface.ShowDatabaseError(6, ex.Message);
                System.Environment.Exit(1);
            }
        }

        internal static void PerformAudit()
        {
            string outputFile = @$"{AuditProperties.outputFilePath}{AuditProperties.auditType} - {AuditProperties.serverGroup.ToUpper()} RESULTS.xlsx";
            using (SqlConnection dbConn = new SqlConnection(DBConnectionString))
            {
                if (dbConn.State == System.Data.ConnectionState.Closed)
                {
                    dbConn.Open();
                }
                var wb = new XLWorkbook();
                DataTable dt = new DataTable();

                //Name of worksheet when saved
                dt.TableName = $"{AuditProperties.dbAuditPrefix.ToUpper()} - {AuditProperties.serverGroup.ToUpper()} AUDIT";
                AuditProperties.servers.Remove(AuditProperties.productionServer);

                // Loop over non-production servers and compare production server > non production for missing servers.
                foreach (string server in AuditProperties.servers)
                {
                    UserInterface.WriteToConsole($"Searching missing values {AuditProperties.productionServer} > {server}...");
                    string query = @$"SELECT 'Missing Property' AS Issue, T1.Name AS 'Property Name', '{AuditProperties.productionServer}' AS 'Primary Instance', '{server}' AS 'Secondary Instance', T1.Value AS 'Production Instance Value', T1.Value AS 'Primary Instance Value', ISNULL(T2.Value, '') AS 'Secondary Instance Value', T1.Type AS 'Type', T1.Application AS 'Application', T1.Description AS 'Description' FROM {AuditProperties.productionServer} T1 LEFT JOIN {server} T2 ON T1.name = T2.name LEFT JOIN [{AuditProperties.exemptionsDatabase}].[dbo].[{AuditProperties.exemptionsTable}] e on T1.Name=e.Name WHERE T2.name IS NULL AND (e.Name IS NULL OR e.Audit_Type != '{AuditProperties.dbAuditPrefix}')";
                    SqlCommand sql = new SqlCommand(query, dbConn);
                    SqlDataAdapter sda = new SqlDataAdapter(sql);
                    sda.Fill(dt);
                }

                // Loop over each non-production server
                foreach (string server in AuditProperties.servers)
                {
                    //  Compare non-prod > prod for missing values
                    if (server != AuditProperties.productionServer)
                    {
                        UserInterface.WriteToConsole($"Searching missing values {server} > {AuditProperties.productionServer}...");
                        string query = @$"SELECT 'Missing Property' AS Issue, T1.Name AS 'Property Name', '{server}' AS 'Primary Instance', '{AuditProperties.productionServer}' AS 'Secondary Instance', T1.Value AS 'Production Instance Value', T1.Value AS 'Primary Instance Value', ISNULL(T2.Value, '') AS 'Secondary Instance Value', T1.Type AS 'Type', T1.Application AS 'Application', T1.Description AS 'Description' FROM {server} T1 LEFT JOIN {AuditProperties.productionServer} T2 ON T1.name = T2.name LEFT JOIN [{AuditProperties.exemptionsDatabase}].[dbo].[{AuditProperties.exemptionsTable}] e on T1.Name=e.Name WHERE T2.name IS NULL AND (e.Name IS NULL OR e.Audit_Type != '{AuditProperties.dbAuditPrefix}')";
                        SqlCommand sql = new SqlCommand(query, dbConn);
                        SqlDataAdapter sda = new SqlDataAdapter(sql);
                        sda.Fill(dt);
                    }

                    // Second loop of servers list
                    foreach (string innerServer in AuditProperties.servers)
                    {
                        // Compare non-prod > other non-prod for missing values
                        if (server != innerServer && server != AuditProperties.productionServer)
                        {
                            UserInterface.WriteToConsole($"Searching missing values {server} > {innerServer}...");
                            string query = @$"SELECT 'Missing Property' AS Issue, T1.Name AS 'Property Name', '{server}' AS 'Primary Instance', '{innerServer}' AS 'Secondary Instance', T1.Value AS 'Production Instance Value', T1.Value AS 'Primary Instance Value', ISNULL(T2.Value, '') AS 'Secondary Instance Value', T1.Type AS 'Type', T1.Application AS 'Application', T1.Description AS 'Description' FROM {server} T1 LEFT JOIN {innerServer} T2 ON T1.name = T2.name LEFT JOIN [{AuditProperties.exemptionsDatabase}].[dbo].[{AuditProperties.exemptionsTable}] e on T1.Name=e.Name WHERE T2.name IS NULL AND (e.Name IS NULL OR e.Audit_Type != '{AuditProperties.dbAuditPrefix}')";
                            SqlCommand sql = new SqlCommand(query, dbConn);
                            SqlDataAdapter sda = new SqlDataAdapter(sql);
                            sda.Fill(dt);
                        }
                    }
                }

                // Create a dictionary to store completed value compares
                Dictionary<string, string> mismatchDone = new Dictionary<string, string>();

                //Search prod server for mismatches against non-prod servers
                foreach (string server in AuditProperties.servers)
                {
                    UserInterface.WriteToConsole($"Searching value mismatch {AuditProperties.productionServer} > {server}......");
                    string query = @$"SELECT 'Value Mismatch' AS Issue, T1.Name AS 'Property Name', '{AuditProperties.productionServer}' AS 'Primary Instance', '{server}' AS 'Secondary Instance', ISNULL(T3.Value, '') AS 'Production Instance Value', T1.Value AS 'Primary Instance Value', T2.Value AS 'Secondary Instance Value', T1.Type AS 'Type', T1.Application AS 'Application', T1.Description AS 'Description' FROM {AuditProperties.productionServer} T1 LEFT JOIN {server} T2 ON T1.name = T2.name LEFT JOIN {AuditProperties.productionServer} T3 ON T1.Name = T3.Name LEFT JOIN [{AuditProperties.exemptionsDatabase}].[dbo].[{AuditProperties.exemptionsTable}] e on T1.Name=e.Name WHERE T1.Value != T2.Value  AND (e.Name IS NULL OR e.Audit_Type != '{AuditProperties.dbAuditPrefix}')";
                    SqlCommand sql = new SqlCommand(query, dbConn);
                    SqlDataAdapter sda = new SqlDataAdapter(sql);
                    sda.Fill(dt);
                }
                mismatchDone.Add(AuditProperties.productionServer, "done");

                //Search non-prod servers against each other for mismatch
                foreach (string server in AuditProperties.servers)
                {
                    foreach (string innerServer in AuditProperties.servers)
                    {
                        if (!mismatchDone.ContainsKey(innerServer) && server != innerServer)
                        {
                            UserInterface.WriteToConsole($"Searching value mismatch {server} > {innerServer}...");
                            string query = @$"SELECT 'Value Mismatch' AS Issue, T1.Name AS 'Property Name', '{server}' AS 'Primary Instance', '{innerServer}' AS 'Secondary Instance', ISNULL(T3.Value, '') AS 'Production Instance Value', T1.Value AS 'Primary Instance Value', T2.Value AS 'Secondary Instance Value', T1.Type AS 'Type', T1.Application AS 'Application', T1.Description AS 'Description' FROM {server} T1 LEFT JOIN {innerServer} T2 ON T1.name = T2.name LEFT JOIN {AuditProperties.productionServer} T3 ON T1.Name = T3.Name LEFT JOIN [{AuditProperties.exemptionsDatabase}].[dbo].[{AuditProperties.exemptionsTable}] e on T1.Name=e.Name WHERE T1.Value != T2.Value AND (e.Name IS NULL OR e.Audit_Type != '{AuditProperties.dbAuditPrefix}')";
                            SqlCommand sql = new SqlCommand(query, dbConn);
                            SqlDataAdapter sda = new SqlDataAdapter(sql);
                            sda.Fill(dt);
                        }
                    }
                    mismatchDone.Add(server, "done");
                }
                var ws = wb.Worksheets.Add(dt);
                // Set Column Widths for excel
                foreach (var item in AuditProperties.columnWidths)
                {
                    ws.Column(item.Key).Width = item.Value;
                }
                ws.Rows().Style.Alignment.WrapText = AuditProperties.wrapRows;
                ws.SheetView.FreezeRows(AuditProperties.outputFreezeRows);
                ws.SheetView.FreezeColumns(AuditProperties.outputFreezeCols);
                try
                {
                    wb.SaveAs(outputFile);
                }
                catch (Exception ex)
                {
                    UserInterface.ShowDatabaseError(7, ex.Message);
                    System.Environment.Exit(1);
                }
            }
        }
    }
}