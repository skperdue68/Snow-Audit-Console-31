using ClosedXML.Excel;
using System.Collections.Generic;
using System;
using System.Configuration;
using System.Data;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.Linq;
using System.Text;

namespace SnowAudit
{
    internal static class DatabaseOperations
    {
        private static string dbConnectionString = ConfigurationManager.ConnectionStrings["database"].ConnectionString;
        private static string exemptionsDatabase = AuditProperties.exemptionsDatabase;
        private static string exemptionsTable = AuditProperties.exemptionsTable;
        private static string inputFilePath = AuditProperties.inputFilePath;

        internal static void Initialize()
        {
            CheckDBConnection();
            CheckForExemptionsDB(exemptionsDatabase, exemptionsTable);
            string dbName = AuditProperties.dbAuditPrefix + AuditProperties.dbServer + "audit";
            CheckForAuditDB(dbName);
            dbConnectionString = dbConnectionString + $"Initial Catalog={dbName};";
            RemoveAuditTables(dbName);
        }

        private static void CheckDBConnection()
        {
            using (SqlConnection dbConn = new SqlConnection(dbConnectionString))
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
                    UserInterface.ShowError("Database 1", ex.Message);
                }
                dbConn.Close();
            }
        }

        private static void CheckForExemptionsDB(string exemptionsDatabase, string exemptionsTable)
        {
            using (SqlConnection dbConn = new SqlConnection(dbConnectionString))
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
                    UserInterface.ShowError("Database 2", ex.Message);
                }
                dbConn.Close();
            }
        }

        private static void CheckForAuditDB(string dbName)
        {
            using (SqlConnection dbConn = new SqlConnection(dbConnectionString))
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
                }
                catch (Exception ex)
                {
                    UserInterface.ShowError("Database 3", ex.Message);
                }
                dbConn.Close();
            }
        }

        private static void RemoveAuditTables(string dbName)
        {
            using (SqlConnection dbConn = new SqlConnection(dbConnectionString))
            {
                try
                {
                    if (dbConn.State == System.Data.ConnectionState.Closed)
                    {
                        dbConn.Open();
                    }
                    dbConn.ChangeDatabase(dbName);
                    SqlCommand sql = new SqlCommand("SELECT * FROM INFORMATION_SCHEMA.TABLES", dbConn);
                    SqlDataReader dataReader = sql.ExecuteReader();
                    while (dataReader.Read())
                    {
                        string dropTable = @$"DROP TABLE {dataReader["table_name"].ToString()}";
                        SqlCommand dropTableCmd = new SqlCommand(dropTable, dbConn);
                        dropTableCmd.ExecuteNonQuery();
                    }

                }
                catch (Exception ex)
                {
                    UserInterface.ShowError("Database 4", ex.Message);
                    System.Environment.Exit(1);
                }
                dbConn.Close();
            }
        }


        internal static DataSet ReadExcel(DataSet ds)
        {
            UserInterface.LoggerChangeColors("READING INPUT FILES", ConsoleColor.White, ConsoleColor.Blue);
            UserInterface.LoggerChangeColors("", ConsoleColor.DarkBlue, ConsoleColor.White);
            string dbName = AuditProperties.dbAuditPrefix + AuditProperties.dbServer;
            DataTable dt = new DataTable();

            foreach (string server in AuditProperties.servers)
            {
                dt = GetDataTableFromExcel(@$"{inputFilePath}{server}.xlsx", $"{server}", "", true);
                ds.Tables.Add(dt);
            }
            return ds;

        }

        public static DataTable GetDataTableFromExcel(string filePath, string tableName, string sheetname = "", bool hasHeader = true)
        {
            using (var workbook = new XLWorkbook(filePath))
            {
                IXLWorksheet worksheet;
                if (string.IsNullOrEmpty(sheetname))
                    worksheet = workbook.Worksheets.First();
                else
                    worksheet = workbook.Worksheets.FirstOrDefault(x => x.Name == sheetname);

                var rangeRowFirst = worksheet.FirstRowUsed().RowNumber();
                var rangeRowLast = worksheet.LastRowUsed().RowNumber();
                var rangeColFirst = worksheet.FirstColumnUsed().ColumnNumber();
                var rangeColLast = worksheet.LastColumnUsed().ColumnNumber();

                DataTable tbl = new DataTable();
                tbl.TableName = tableName;
                for (int col = rangeColFirst; col <= rangeColLast; col++)
                    tbl.Columns.Add(hasHeader ? worksheet.FirstRowUsed().Cell(col).Value.ToString() : $"Column {col}");
                UserInterface.Logger($"Reading data for {tableName}...", false);
                rangeRowFirst = rangeRowFirst + (hasHeader ? 1 : 0);
                var colCount = rangeColLast - rangeColFirst;
                for (int rowNum = rangeRowFirst; rowNum <= rangeRowLast; rowNum++)
                {
                    List<string> colValues = new List<string>();
                    for (int col = 1; col <= colCount; col++)
                    {
                        colValues.Add(worksheet.Row(rowNum).Cell(col).Value.ToString());
                    }
                    tbl.Rows.Add(colValues.ToArray());
                }
                UserInterface.Logger($"done!");
                return tbl;
            }
        }

        internal static void AutoSqlBulkCopy(DataSet dataSet)
        {
            using (SqlConnection dbConn = new SqlConnection(dbConnectionString))
            {
                if (dbConn.State == System.Data.ConnectionState.Closed)
                {
                    dbConn.Open();
                }
                string dbName = AuditProperties.dbAuditPrefix + AuditProperties.dbServer + "audit";
                dbConn.ChangeDatabase(dbName);
                foreach (System.Data.DataTable dataTable in dataSet.Tables)
                {
                    // checking whether the table selected from the dataset exists in the database or not
                    var checkTableIfExistsCommand = new SqlCommand("IF EXISTS (SELECT 1 FROM sysobjects WHERE name =  '" + dataTable.TableName + "') SELECT 1 ELSE SELECT 0", dbConn);
                    var exists = checkTableIfExistsCommand.ExecuteScalar().ToString().Equals("1");

                    // if table does not exist
                    if (!exists)
                    {
                        var createTableBuilder = new StringBuilder("CREATE TABLE [" + dataTable.TableName + "]");
                        createTableBuilder.AppendLine("(");

                        // selecting each column of the datatable to create a table in the database
                        foreach (DataColumn dc in dataTable.Columns)
                        {
                            createTableBuilder.AppendLine("  [" + dc.ColumnName + "] VARCHAR(MAX),");
                        }

                        createTableBuilder.Remove(createTableBuilder.Length - 1, 1);
                        createTableBuilder.AppendLine(")");

                        var createTableCommand = new SqlCommand(createTableBuilder.ToString(), dbConn);
                        createTableCommand.ExecuteNonQuery();
                    }

                    // if table exists, just copy the data to the destination table in the database
                    // copying the data from datatable to database table
                    using (var bulkCopy = new SqlBulkCopy(dbConn))
                    {
                        bulkCopy.DestinationTableName = dataTable.TableName;
                        bulkCopy.WriteToServer(dataTable);
                    }
                }
            }
        }

        internal static void PerformAudit()
        {
            UserInterface.LoggerChangeColors("PERFORMING AUDIT", ConsoleColor.White, ConsoleColor.Blue);
            UserInterface.LoggerChangeColors("", ConsoleColor.DarkBlue, ConsoleColor.White);
            string outputFile = @$"{AuditProperties.outputFilePath}{AuditProperties.auditType} - {AuditProperties.serverGroup.ToUpper()} RESULTS.xlsx";
            using (SqlConnection dbConn = new SqlConnection(dbConnectionString))
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
                    UserInterface.Logger($"Searching missing values {AuditProperties.productionServer} > {server}...", false);
                    string query = @$"SELECT 'Missing Property' AS Issue, T1.Name AS 'Property Name', '{AuditProperties.productionServer}' AS 'Primary Instance', '{server}' AS 'Secondary Instance', T1.Value AS 'Production Instance Value', T1.Value AS 'Primary Instance Value', ISNULL(T2.Value, '') AS 'Secondary Instance Value', T1.Type AS 'Type', T1.Application AS 'Application', T1.Description AS 'Description' FROM {AuditProperties.productionServer} T1 LEFT JOIN {server} T2 ON T1.name = T2.name LEFT JOIN [{AuditProperties.exemptionsDatabase}].[dbo].[{AuditProperties.exemptionsTable}] e on T1.Name=e.Name WHERE T2.name IS NULL AND (e.Name IS NULL OR e.Audit_Type != '{AuditProperties.dbAuditPrefix}')";
                    SqlCommand sql = new SqlCommand(query, dbConn);
                    SqlDataAdapter sda = new SqlDataAdapter(sql);
                    sda.Fill(dt);
                    UserInterface.Logger("done!");
                }

                // Loop over each non-production server
                foreach (string server in AuditProperties.servers)
                {
                    //  Compare non-prod > prod for missing values
                    if (server != AuditProperties.productionServer)
                    {
                        UserInterface.Logger($"Searching missing values {server} > {AuditProperties.productionServer}...", false);
                        string query = @$"SELECT 'Missing Property' AS Issue, T1.Name AS 'Property Name', '{server}' AS 'Primary Instance', '{AuditProperties.productionServer}' AS 'Secondary Instance', T1.Value AS 'Production Instance Value', T1.Value AS 'Primary Instance Value', ISNULL(T2.Value, '') AS 'Secondary Instance Value', T1.Type AS 'Type', T1.Application AS 'Application', T1.Description AS 'Description' FROM {server} T1 LEFT JOIN {AuditProperties.productionServer} T2 ON T1.name = T2.name LEFT JOIN [{AuditProperties.exemptionsDatabase}].[dbo].[{AuditProperties.exemptionsTable}] e on T1.Name=e.Name WHERE T2.name IS NULL AND (e.Name IS NULL OR e.Audit_Type != '{AuditProperties.dbAuditPrefix}')";
                        SqlCommand sql = new SqlCommand(query, dbConn);
                        SqlDataAdapter sda = new SqlDataAdapter(sql);
                        sda.Fill(dt);
                        UserInterface.Logger("done!");
                    }

                    // Second loop of servers list
                    foreach (string innerServer in AuditProperties.servers)
                    {
                        // Compare non-prod > other non-prod for missing values
                        if (server != innerServer && server != AuditProperties.productionServer)
                        {
                            UserInterface.Logger($"Searching missing values {server} > {innerServer}...", false);
                            string query = @$"SELECT 'Missing Property' AS Issue, T1.Name AS 'Property Name', '{server}' AS 'Primary Instance', '{innerServer}' AS 'Secondary Instance', T1.Value AS 'Production Instance Value', T1.Value AS 'Primary Instance Value', ISNULL(T2.Value, '') AS 'Secondary Instance Value', T1.Type AS 'Type', T1.Application AS 'Application', T1.Description AS 'Description' FROM {server} T1 LEFT JOIN {innerServer} T2 ON T1.name = T2.name LEFT JOIN [{AuditProperties.exemptionsDatabase}].[dbo].[{AuditProperties.exemptionsTable}] e on T1.Name=e.Name WHERE T2.name IS NULL AND (e.Name IS NULL OR e.Audit_Type != '{AuditProperties.dbAuditPrefix}')";
                            SqlCommand sql = new SqlCommand(query, dbConn);
                            SqlDataAdapter sda = new SqlDataAdapter(sql);
                            sda.Fill(dt);
                            UserInterface.Logger("done!");
                        }
                    }
                }

                // Create a dictionary to store completed value compares
                Dictionary<string, string> mismatchDone = new Dictionary<string, string>();

                //Search prod server for mismatches against non-prod servers
                foreach (string server in AuditProperties.servers)
                {
                    UserInterface.Logger($"Searching value mismatch {AuditProperties.productionServer} > {server}......", false);
                    string query = @$"SELECT 'Value Mismatch' AS Issue, T1.Name AS 'Property Name', '{AuditProperties.productionServer}' AS 'Primary Instance', '{server}' AS 'Secondary Instance', ISNULL(T3.Value, '') AS 'Production Instance Value', T1.Value AS 'Primary Instance Value', T2.Value AS 'Secondary Instance Value', T1.Type AS 'Type', T1.Application AS 'Application', T1.Description AS 'Description' FROM {AuditProperties.productionServer} T1 LEFT JOIN {server} T2 ON T1.name = T2.name LEFT JOIN {AuditProperties.productionServer} T3 ON T1.Name = T3.Name LEFT JOIN [{AuditProperties.exemptionsDatabase}].[dbo].[{AuditProperties.exemptionsTable}] e on T1.Name=e.Name WHERE T1.Value != T2.Value  AND (e.Name IS NULL OR e.Audit_Type != '{AuditProperties.dbAuditPrefix}')";
                    SqlCommand sql = new SqlCommand(query, dbConn);
                    SqlDataAdapter sda = new SqlDataAdapter(sql);
                    sda.Fill(dt);
                    UserInterface.Logger("done!");
                }
                mismatchDone.Add(AuditProperties.productionServer, "done");

                //Search non-prod servers against each other for mismatch
                foreach (string server in AuditProperties.servers)
                {
                    foreach (string innerServer in AuditProperties.servers)
                    {
                        if (!mismatchDone.ContainsKey(innerServer) && server != innerServer)
                        {
                            UserInterface.Logger($"Searching value mismatch {server} > {innerServer}...", false);
                            string query = @$"SELECT 'Value Mismatch' AS Issue, T1.Name AS 'Property Name', '{server}' AS 'Primary Instance', '{innerServer}' AS 'Secondary Instance', ISNULL(T3.Value, '') AS 'Production Instance Value', T1.Value AS 'Primary Instance Value', T2.Value AS 'Secondary Instance Value', T1.Type AS 'Type', T1.Application AS 'Application', T1.Description AS 'Description' FROM {server} T1 LEFT JOIN {innerServer} T2 ON T1.name = T2.name LEFT JOIN {AuditProperties.productionServer} T3 ON T1.Name = T3.Name LEFT JOIN [{AuditProperties.exemptionsDatabase}].[dbo].[{AuditProperties.exemptionsTable}] e on T1.Name=e.Name WHERE T1.Value != T2.Value AND (e.Name IS NULL OR e.Audit_Type != '{AuditProperties.dbAuditPrefix}')";
                            SqlCommand sql = new SqlCommand(query, dbConn);
                            SqlDataAdapter sda = new SqlDataAdapter(sql);
                            sda.Fill(dt);
                            UserInterface.Logger("done!");
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
                    UserInterface.ShowError("Database 7", ex.Message);
                    System.Environment.Exit(1);
                }
            }
        }
    }
}