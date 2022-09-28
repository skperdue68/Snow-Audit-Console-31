using MenuBuilder;
using System.Linq;
using System;
using System.Text.RegularExpressions;

namespace SnowAudit
{
    public static class UserInterface
    {
        static string inputFilePath = AuditProperties.inputFilePath;
        static string outputFilePath = AuditProperties.outputFilePath;

        public static void DisplayInfo()
        {
            ClearConsole();
            Logger($"ServiceNOW Audits was created to assist in automating completing properties audits for ServiceNOW.");
            Logger($"\r\nExport the necessary property list(s) from ServiceNOW, saving them as Excel Files.");
            Logger($"\r\nSave files as the name of the instance you are exporting (i.e. attbdas, attbdasdev, attfedgov1, attfedgov1test, etc) in '<color>{inputFilePath}</color>.'", ConsoleColor.DarkYellow);
            Logger($"\r\nThe application will automatically attempt to determine which file(s) are appropriate based on the server group selected.");
            Logger($"\r\nUpon completion of the audit your results will be found in '<color>{outputFilePath}</color>'.", ConsoleColor.DarkYellow);
            Pause();
        }

        public static int SelectAction()
        {
            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.ForegroundColor = ConsoleColor.White;
            ClearConsole();
            LoggerChangeColors("SELECT ACTION                       ", ConsoleColor.White, ConsoleColor.Blue);
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
                    case 3:
                        {
                            Environment.Exit(0);
                            break;
                        }
                }
            }
            return selection;
        }

        public static int SelectAuditType()
        {
            ClearConsole();
            LoggerChangeColors("PERFORM AUDIT > SELECT AUDIT TYPE   ", ConsoleColor.White, ConsoleColor.Blue);
            string[] menuOptions = { "SYSTEM PROPERTIES", "BACK TO ACTIONS" };
            var newMenu = new Menu(menuOptions, 2, 3);
            newMenu.ModifyMenuLeftJustified();
            newMenu.ResetCursorVisible();
            int selection = 0;
            while (selection == 0)
            {
                selection = newMenu.RunMenu();
                AuditProperties.auditType = menuOptions[selection - 1].Trim();
            }
            return selection;
        }

        public static int SelectServerGroup()
        {
            ClearConsole();
            LoggerChangeColors($"PERFORM AUDIT > {AuditProperties.auditType} > SELECT INSTANCE", ConsoleColor.White, ConsoleColor.Blue);
            string[] menuOptions = { "BDAS", "FUSION", "FEDGOV", "BACK TO AUDIT TYPE" };
            var newMenu = new Menu(menuOptions, 2, 3);
            newMenu.ModifyMenuLeftJustified();
            newMenu.ResetCursorVisible();
            int selection = 0;
            while (selection == 0)
            {
                selection = newMenu.RunMenu();
            }
            ClearConsole();
            if (selection < newMenu.length) AuditProperties.serverGroup = menuOptions[selection - 1].Trim();
            return selection;
        }

        public static void ShowError(string errorCode, string errorText)
        {
            Console.ForegroundColor = ConsoleColor.Red;
            switch (errorCode)
            {
                case "File 1":
                    Logger($"ERROR: An input file for production server {inputFilePath}{AuditProperties.productionServer}.xlsx not found.", ConsoleColor.DarkRed);
                    break;
                case "File 2":
                    Logger("ERROR: A required input file for at least one non-production server to compare was not found.", ConsoleColor.DarkRed);
                    break;
                case "Database 1":
                    Logger("ERROR: Unable to connect to to the database server.", ConsoleColor.DarkRed);
                    break;
                case "Database 2":
                    Logger("ERROR: There is an issue with the exemtion table.", ConsoleColor.DarkRed);
                    break;
                case "Database 3":
                    Logger("ERROR: There is an issue verifying or creating the audit database tables.", ConsoleColor.DarkRed);
                    break;
                case "Database 4":
                    Logger("ERROR: Unable to remove existing audit server data tables.", ConsoleColor.DarkRed);
                    Logger(errorText);
                    break;
                case "Database 5":
                    Logger("ERROR: Unable to create audit server data tables.", ConsoleColor.DarkRed);
                    break;
                case "Database 6":
                    Logger("ERROR: Unable to import server data to database.", ConsoleColor.DarkRed);
                    break;
                case "Database 7":
                    Logger("ERROR: Unable to save output file.", ConsoleColor.DarkRed);
                    break;
            }
            Environment.Exit(0);
        }

        //Clear Console and optionally set background and foreground colors.
        internal static void ClearConsole(ConsoleColor foregroundColor = ConsoleColor.White, ConsoleColor backgroundColor = ConsoleColor.DarkBlue)
        {
            Console.ForegroundColor = ConsoleColor.White;
            Console.BackgroundColor = ConsoleColor.DarkBlue;
            Console.Clear();
        }


        //Display a line while optionally changing background and/or foreground
        internal static void LoggerChangeColors(string message, ConsoleColor BackgroundColor = ConsoleColor.DarkBlue, ConsoleColor ForegroundColor = ConsoleColor.White)
        {
            Console.ForegroundColor = ForegroundColor;
            Console.BackgroundColor = BackgroundColor;
            WordWrapper.WordWrapper.Wrap(message);

        }

        //Display a line with a color highlighted portion.
        internal static void Logger(string message, ConsoleColor newColor = ConsoleColor.White)
        {

            // Set defaults
            ConsoleColor originalColor = Console.ForegroundColor;
            bool isMessageWrapped = false;
            string messageNoWrap = String.Empty;

            //Split the message into an array of strings based on <color></color> tags
            var messagePieces = Regex.Split(message, @"(<color>[\s\S]+?<\/color>)").Where(l => l != string.Empty).ToArray();

            // The entire message either has no <color></color> tags or the entire thing wrapped in color, either way this can be wrapped as normal
            if (messagePieces.Length == 1)
            {
                foreach (var messagePiece in messagePieces)
                {
                    isMessageWrapped = Regex.Match(messagePiece, @"(<color>[\s\S]+?<\/color>)").Success;

                    // Change Color
                    if (isMessageWrapped)
                    {
                        Console.ForegroundColor = newColor;
                        messageNoWrap = Regex.Match(messagePiece, @"(?<=<color>)(.*?)(?=<\/color>)").Groups[1].Value;
                    }
                    else
                    {
                        messageNoWrap = messagePiece;
                    }
                    WordWrapper.WordWrapper.Wrap(messageNoWrap);
                    if (isMessageWrapped)
                    {
                        Console.ForegroundColor = originalColor;
                    }
                }
            }
            else
            {
                // Start by creating string removing matching tags, keeping everything else.  Will I even need this?
                string messagePiece = String.Empty;
                int consolePosition = 0;

                for (int i = 0; i < messagePieces.Length; i++)
                {
                    messagePiece = messagePieces[i];
                    int windowWidth = Console.WindowWidth - 1;
                    isMessageWrapped = Regex.Match(messagePiece, @"(<color>[\s\S]+?<\/color>)").Success;

                    if (isMessageWrapped)
                    {
                        Console.ForegroundColor = newColor;
                        messageNoWrap = Regex.Match(messagePiece, @"(?<=<color>)(.*?)(?=<\/color>)").Groups[1].Value;
                    }
                    else
                    {
                        messageNoWrap = messagePiece;
                    }
                    if (isMessageWrapped)
                    {
                        Console.ForegroundColor = newColor;
                    }
                    var messageWords = messageNoWrap.Split(' ');
                    messageWords = Regex.Split(messageNoWrap, @"(?<=\s+)");

                    for (int j = 0; j < messageWords.Length; j++)
                    {
                        if (consolePosition + messageWords[j].Length <= windowWidth)
                        {
                            consolePosition = consolePosition + messageWords[j].Length;
                            Console.Write(messageWords[j]);
                        }
                        else
                        {
                            Console.WriteLine();
                            consolePosition = 0;
                            j--;
                        }
                    }
                    if (isMessageWrapped)
                    {
                        Console.ForegroundColor = originalColor;
                    }
                }
                Console.WriteLine();
            }
        }

        internal static void Pause()
        {
            Console.WriteLine("\r\n\r\nPRESS ANY KEY TO CONTINUE");
            Console.ReadKey();
        }

        internal static void NotImplemented()
        {
            Console.WriteLine("This functionality has not been completed.");
            Pause();
            ClearConsole();
        }
    }
}