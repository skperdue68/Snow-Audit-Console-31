using MenuBuilder;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Net.WebSockets;
using System.Runtime;
using System.Text;
using System.Threading.Tasks;
using System.Threading.Tasks.Sources;

namespace SnowAudit
{
    public static class UserInterface
    {
        public static void DisplayInfo()
        {
            ClearConsole();
            Console.WriteLine("This application will take excel .xlsx files exported from SNOW and will compare them for missing or different values and provide you an excel spreadsheet as results.");
            Console.WriteLine();
            Console.WriteLine($"The excel files should be named the same as the instance you pulled then from (i.e. attbdasdev2.xlsx, attfusionbeta.xlsx, fedgov1dev2.xlsx, etc) and placed in {AuditProperties.inputFilePath}");
            Console.WriteLine();
            Console.WriteLine("Press Enter to begin...");
            Console.ReadLine();
        }

        public static int SelectAction()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.White;
            Console.BackgroundColor = ConsoleColor.Black;
            Console.WriteLine("SELECT ACTION");
            string[] menuOptions = { "PERFORM AUDIT", "AUDIT EXCEPTIONS MAINTENANCE", "EXIT" };
            var newMenu = new Menu(menuOptions, 2, 3);
            newMenu.ModifyMenuLeftJustified();
            newMenu.ResetCursorVisible();
            int selection = 0;
            while (selection == 0)
            {
                selection = newMenu.RunMenu();
                switch (selection)
                {
                    case 1:
                        {
                            break;
                        }
                    case 2:
                        {
                            break;
                        }
                    case 3:
                        {
                            Environment.Exit(0);
                            break;
                        }
                }
            }
            return selection;
        }

        public static int AuditSelectAuditType()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.White;
            Console.BackgroundColor = ConsoleColor.Black;
            Console.WriteLine("PERFORM AUDIT > SELECT AUDIT TYPE");
            string[] menuOptions = { "SYSTEM PROPERTIES", "BACK TO ACTIONS" };
            var newMenu = new Menu(menuOptions, 2, 3);
            newMenu.ModifyMenuLeftJustified();
            newMenu.ResetCursorVisible();
            int selection = 0;
            while (selection == 0)
            {
                selection = newMenu.RunMenu();
                switch (selection)
                {
                    case 1:
                        {
                            break;
                        }
                    case 2:
                        {
                            break;
                        }
                }
                AuditProperties.auditType = menuOptions[selection - 1].Trim();
            }
            return selection;
        }

        public static int SelectInstanceType()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.White;
            Console.BackgroundColor = ConsoleColor.Black;
            Console.WriteLine($"PERFORM AUDIT > {AuditProperties.auditType} > SELECT INSTANCE");
            string[] menuOptions = { "BDAS", "FUSION", "FEDGOV", "BACK TO AUDIT TYPE" };
            var newMenu = new Menu(menuOptions, 2, 3);
            newMenu.ModifyMenuLeftJustified();
            newMenu.ResetCursorVisible();
            int selection = 0;
            while (selection == 0)
            {
                selection = newMenu.RunMenu();
            }
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.White;
            Console.BackgroundColor = ConsoleColor.Black;
            if (selection < newMenu.length) AuditProperties.serverGroup = menuOptions[selection - 1].Trim();
            return selection;
        }

        public static void ShowInputFileError(int errorCode)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            switch (errorCode)
            {
                case 1:

                    Console.WriteLine($"An input file for production server {AuditProperties.inputFilePath}{AuditProperties.productionServer}.xlsx not found.");
                    break;
                case 2:
                    Console.WriteLine("A required input file for at least one non-production server to compare not found.");
                    break;
                default:
                    Console.WriteLine("An unknown file error occurred.");
                    break;
            }
            Environment.Exit(0);
        }

        public static void ShowDatabaseError(int code, string message)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            switch (code)
            {
                case 1:
                    {
                        Console.WriteLine("Unable to connect to to the database server.");
                        break;
                    }
                case 2:
                    {
                        Console.WriteLine("An errpr occcurred with exemption tables.");
                        break;
                    }
                case 3:
                    {
                        Console.WriteLine("An error occurred dropping old audit database.");
                        break;
                    }
                case 4:
                    {
                        Console.WriteLine("An error occurred creating database for audit.");
                        break;
                    }
                case 5:
                    {
                        Console.WriteLine("An error occurred creating database tables for audit.");
                        break;
                    }
                case 6:
                    {
                        Console.WriteLine("An error occurred importing files to database.");
                        break;
                    }
                case 7:
                    {
                        Console.WriteLine("An error saving output.");
                        break;
                    }
                default:
                    {
                        Console.WriteLine("An unknown database occurred with database operations");
                        break;
                    }
            }
            Console.WriteLine(message);
            Environment.Exit(0);
        }

        internal static void ClearConsole()
        {
            Console.Clear();
            Console.ForegroundColor = ConsoleColor.White;
            Console.BackgroundColor = ConsoleColor.Black;
        }

        internal static void WriteToConsole(string message)
        {
            Console.WriteLine(message);
        }

        internal static void NotImplemented()
        {
            Console.WriteLine("This has not been implemented yet");
            Console.WriteLine("Press Return to go back");
            Console.ReadLine();
            ClearConsole();
        }
    }
}