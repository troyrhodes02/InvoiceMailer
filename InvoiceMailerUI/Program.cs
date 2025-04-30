using System;
using System.CommandLine;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using Spectre.Console;
using Spectre.Console.Cli;
using Microsoft.Extensions.Configuration;

namespace InvoiceMailerUI
{
    public class Program
    {
        private static readonly InvoiceMailerController _controller = new InvoiceMailerController();
        private static readonly List<(string Message, InvoiceMailerController.LogLevel Level)> _logEntries = new();

        public static async Task<int> Main(string[] args)
        {
            // Setup the controller log callback
            _controller.SetLogCallback(LogMessageCallback);
            
            // Run the setup wizard first to ensure configuration is complete
            if (!await _controller.RunSetupWizard())
            {
                // If setup wizard fails, exit the application
                return 1;
            }
            
            // Ensure required folders exist for production use
            _controller.VerifyConfiguration();

            // Display the application header
            AnsiConsole.Clear();
            AnsiConsole.Write(new FigletText("Invoice Mailer").Color(Color.Orange1));
            AnsiConsole.WriteLine();
            AnsiConsole.MarkupLine("[blue]Production Email Sender[/] - [yellow]v1.0[/]");
            AnsiConsole.WriteLine();

            return await ShowMainMenu();
        }

        private static async Task<int> ShowMainMenu()
        {
            while (true)
            {
                var choice = AnsiConsole.Prompt(
                    new SelectionPrompt<string>()
                        .Title("What would you like to do?")
                        .PageSize(10)
                        .AddChoices(
                            "Send Invoices",
                            "Dry Run (Test Mode)",
                            "Health Check",
                            "Configure Settings",
                            "View Log",
                            "Help",
                            "Exit"));

                switch (choice)
                {
                    case "Send Invoices":
                        await RunInvoiceSender(false);
                        break;
                    case "Dry Run (Test Mode)":
                        await RunInvoiceSender(true);
                        break;
                    case "Health Check":
                        await RunHealthCheck();
                        break;
                    case "Configure Settings":
                        await ConfigureSettings();
                        break;
                    case "View Log":
                        ViewLogs();
                        break;
                    case "Help":
                        ShowMainHelp();
                        break;
                    case "Exit":
                        AnsiConsole.Clear();
                        AnsiConsole.MarkupLine("[green]Thank you for using InvoiceMailer![/]");
                        AnsiConsole.WriteLine();
                        return 0;
                }
            }
        }

        private static void ViewLogs()
        {
            AnsiConsole.Clear();
            AnsiConsole.Write(new Rule("[yellow]System Log[/]").RuleStyle("grey").Centered());
            AnsiConsole.WriteLine();

            if (_logEntries.Count == 0)
            {
                AnsiConsole.MarkupLine("[grey]No log entries available.[/]");
            }
            else
            {
                // Create a scrollable panel to display logs
                var logPanel = new Panel(GetFormattedLogs())
                {
                    Header = new PanelHeader("Application Log"),
                    Border = BoxBorder.Rounded,
                    Expand = true,
                    Padding = new Padding(1, 1, 1, 1)
                };
                
                AnsiConsole.Write(logPanel);
            }

            AnsiConsole.WriteLine();
            Console.WriteLine("Press any key to return to the main menu...");
            Console.ReadKey();
            AnsiConsole.Clear();
        }

        private static string GetFormattedLogs()
        {
            var formattedLog = new System.Text.StringBuilder();
            
            foreach (var entry in _logEntries)
            {
                string colorCode = entry.Level switch
                {
                    InvoiceMailerController.LogLevel.Info => "blue",
                    InvoiceMailerController.LogLevel.Success => "green",
                    InvoiceMailerController.LogLevel.Warning => "yellow",
                    InvoiceMailerController.LogLevel.Error => "red",
                    _ => "white"
                };
                
                formattedLog.AppendLine($"[{colorCode}]{entry.Message}[/]");
            }
            
            return formattedLog.ToString();
        }

        private static void LogMessageCallback(string message, InvoiceMailerController.LogLevel level)
        {
            // Add the message to our log entries
            _logEntries.Add((message, level));
            
            // If we're in a status context, we can't write to console directly
            // The messages will be shown later in the log viewer
        }

        private static void ShowMainHelp()
        {
            AnsiConsole.Clear();
            AnsiConsole.Write(new Rule("[yellow]Help: Main Menu[/]").RuleStyle("grey").Centered());
            
            AnsiConsole.WriteLine();
            AnsiConsole.Write(new Panel(
                "The InvoiceMailer application processes invoice files and sends emails to recipients.\n\n" +
                "[bold]Menu Options:[/]\n" +
                "• [green]Send Invoices[/]: Process invoice files and send actual emails to recipients.\n" +
                "• [green]Dry Run[/]: Simulate the email sending process without actually sending emails.\n" +
                "• [green]Health Check[/]: Verify that all required components are properly configured.\n" +
                "• [green]Configure Settings[/]: View and modify application settings.\n" +
                "• [green]View Log[/]: Display the application log with color-coded messages.\n" +
                "• [green]Help[/]: Display this help information.\n" +
                "• [green]Exit[/]: Close the application."
            ).Border(BoxBorder.Rounded).Padding(1, 1));
            
            Console.WriteLine("\nPress any key to return to the main menu...");
            Console.ReadKey();
            AnsiConsole.Clear();
        }

        private static async Task RunInvoiceSender(bool dryRun)
        {
            bool goBack = false;
            
            while (!goBack)
            {
                AnsiConsole.Clear();
                AnsiConsole.Write(new Rule(dryRun ? "[yellow]Dry Run Mode[/]" : "[yellow]Send Invoices[/]").RuleStyle("grey").Centered());
                AnsiConsole.WriteLine();
                
                var choice = AnsiConsole.Prompt(
                    new SelectionPrompt<string>()
                        .Title(dryRun ? "Dry Run Options:" : "Send Invoice Options:")
                        .PageSize(10)
                        .AddChoices(
                            "Start Process",
                            "Help",
                            "Back to Main Menu"));
                
                switch (choice)
                {
                    case "Start Process":
                        await ProcessInvoices(dryRun);
                        break;
                    case "Help":
                        ShowInvoiceSenderHelp(dryRun);
                        break;
                    case "Back to Main Menu":
                        goBack = true;
                        break;
                }
            }
            
            AnsiConsole.Clear();
        }
        
        private static void ShowInvoiceSenderHelp(bool dryRun)
        {
            AnsiConsole.Clear();
            AnsiConsole.Write(new Rule(dryRun ? "[yellow]Help: Dry Run Mode[/]" : "[yellow]Help: Send Invoices[/]").RuleStyle("grey").Centered());
            
            AnsiConsole.WriteLine();
            AnsiConsole.Write(new Panel(
                $"[bold]{(dryRun ? "Dry Run Mode" : "Send Invoices")}[/]\n\n" +
                $"This feature {(dryRun ? "simulates" : "executes")} the invoice processing and email sending flow.\n\n" +
                "[bold]How it works:[/]\n" +
                "1. You'll be asked to specify the invoices folder path (or use the default)\n" +
                "2. You'll be asked to specify the recipients CSV file path (or use the default)\n" +
                "3. The application will scan for invoice files and match them with recipients\n" +
                $"4. {(dryRun ? "Emails will be simulated but not actually sent" : "Emails will be sent to the specified recipients")}\n\n" +
                "[bold]Tips:[/]\n" +
                "• Leave the path fields empty to use the default paths\n" +
                "• Make sure your invoice files follow the naming pattern (e.g., INV12345)\n" +
                "• The recipients CSV should have 'InvoiceKey,Email' format\n\n" +
                "[bold]Security Note:[/]\n" +
                "• All credentials are loaded securely from appsettings.json\n" +
                "• Authentication is handled securely through interactive login\n" +
                "• Emails are sent using your authenticated Microsoft 365 account"
            ).Border(BoxBorder.Rounded).Padding(1, 1));
            
            Console.WriteLine("\nPress any key to return...");
            Console.ReadKey();
            AnsiConsole.Clear();
        }
        
        /// <summary>
        /// Process invoices using the sequential phase execution pattern
        /// </summary>
        private static async Task ProcessInvoices(bool dryRun)
        {
            // PHASE 1: SETUP & AUTH - Collect input and authenticate BEFORE spinner
            AnsiConsole.Clear();
            AnsiConsole.Write(new Rule(dryRun ? "[yellow]Dry Run Mode - Setup[/]" : "[yellow]Send Invoices - Setup[/]").RuleStyle("grey").Centered());
            AnsiConsole.WriteLine();
            
            // Get user input for paths
            var invoicesFolder = AnsiConsole.Prompt(
                new TextPrompt<string>("Enter invoices folder path (or press Enter for default):")
                    .AllowEmpty());

            var recipientsFile = AnsiConsole.Prompt(
                new TextPrompt<string>("Enter recipients file path (or press Enter for default):")
                    .AllowEmpty());

            // Get the default sender email from configuration for informational purposes
            var emailConfig = _controller.GetEmailConfiguration();
            var defaultSenderEmail = emailConfig.GetValueOrDefault("DefaultSenderEmail", "");
            // Note: Actual authenticated user email isn't known until after auth
            var initialSenderInfo = !string.IsNullOrEmpty(defaultSenderEmail) ? defaultSenderEmail : "your authenticated account";
            
            // Ask about sender email override
            string? senderEmail = null;
            var overrideSender = AnsiConsole.Confirm($"Emails will default to {initialSenderInfo}. Would you like to specify a different sender for this run?", false);
            
            if (overrideSender)
            {
                senderEmail = AnsiConsole.Prompt(
                    new TextPrompt<string>("Enter the sender email address:")
                        .DefaultValue(defaultSenderEmail) // Suggest default if it exists
                        .Validate(email => 
                        {
                            if (string.IsNullOrWhiteSpace(email))
                                return ValidationResult.Error("Email cannot be empty");
                            if (!email.Contains('@'))
                                return ValidationResult.Error("Invalid email format");
                            return ValidationResult.Success();
                        }));
            }

            // Add confirmation with back option
            AnsiConsole.WriteLine();
            var proceedChoice = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Ready to proceed?")
                    .PageSize(3)
                    .AddChoices(
                        "Proceed",
                        "Back to Previous Screen"
                    ));
            
            if (proceedChoice == "Back to Previous Screen")
            {
                AnsiConsole.Clear();
                return;
            }

            // Confirm before sending real emails
            if (!dryRun)
            {
                var confirmed = AnsiConsole.Confirm("WARNING: This will send REAL emails to recipients. Continue?", false);
                if (!confirmed)
                {
                    AnsiConsole.MarkupLine("[yellow]Operation cancelled by user.[/]");
                    Console.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                    AnsiConsole.Clear();
                    return;
                }
            }
            
            // *** PERFORM AUTHENTICATION HERE (BEFORE SPINNER) ***
            AnsiConsole.WriteLine(); // Separator
            AnsiConsole.MarkupLine("[cyan]Attempting authentication...[/]");
            bool isAuthenticated = await _controller.AuthenticateUser(dryRun); // Use the actual dryRun parameter
            
            if (!isAuthenticated)
            {
                AnsiConsole.MarkupLine("[red]Authentication failed. Cannot proceed with sending emails.[/]");
                Console.WriteLine("Press any key to return to the previous screen...");
                Console.ReadKey();
                AnsiConsole.Clear();
                return;
            }
            AnsiConsole.MarkupLine("[green]Authentication successful.[/]");
            AnsiConsole.WriteLine(); // Separator
            
            // PHASE 2: PROCESSING - After all user input AND authentication is collected
            AnsiConsole.Clear();
            AnsiConsole.Write(new Rule(dryRun ? "[yellow]Dry Run Mode - Processing[/]" : "[yellow]Send Invoices - Processing[/]").RuleStyle("grey").Centered());
            AnsiConsole.WriteLine();
            
            bool success = false;
            
            // Use a status display for the processing phase - NO interactive prompts or auth here
            await AnsiConsole.Status()
                .StartAsync("Processing invoices...", async ctx =>
                {
                    ctx.Spinner(Spinner.Known.Dots);
                    
                    // Only log steps here, no actual operations that require interaction
                    ctx.Status("Scanning for invoices...");
                    ctx.Status("Matching recipients...");
                    ctx.Status(dryRun ? "Simulating email sending..." : "Sending emails...");

                    // Use the controller to run the appropriate process (auth already done)
                    if (dryRun)
                    {
                        // Pass collected info to RunDryRun which now calls RunInvoiceMailer directly
                        success = await _controller.RunDryRun(invoicesFolder, recipientsFile, senderEmail);
                    }
                    else
                    {
                        // Pass collected info to RunRealSend which now calls RunInvoiceMailer directly
                        success = await _controller.RunRealSend(invoicesFolder, recipientsFile, senderEmail);
                    }
                    
                    ctx.Status("Finishing process...");
                });

            // PHASE 3: RESULTS - Display final results
            AnsiConsole.WriteLine();
            if (success)
            {
                AnsiConsole.MarkupLine("[green]Process completed successfully![/]");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]Process completed with issues. See log for details.[/]");
            }
            AnsiConsole.MarkupLine("[blue]Use the 'View Log' option to see detailed information.[/]");
            Console.WriteLine("Press any key to return to the previous screen...");
            Console.ReadKey();
            AnsiConsole.Clear();
        }

        private static async Task RunHealthCheck()
        {
            bool goBack = false;
            
            while (!goBack)
            {
                AnsiConsole.Clear();
                AnsiConsole.Write(new Rule("[yellow]Health Check[/]").RuleStyle("grey").Centered());
                AnsiConsole.WriteLine();
                
                var choice = AnsiConsole.Prompt(
                    new SelectionPrompt<string>()
                        .Title("Health Check Options:")
                        .PageSize(10)
                        .AddChoices(
                            "Start Health Check",
                            "Help",
                            "Back to Main Menu"));
                
                switch (choice)
                {
                    case "Start Health Check":
                        await PerformHealthCheck();
                        break;
                    case "Help":
                        ShowHealthCheckHelp();
                        break;
                    case "Back to Main Menu":
                        goBack = true;
                        break;
                }
            }
            
            AnsiConsole.Clear();
        }
        
        private static void ShowHealthCheckHelp()
        {
            AnsiConsole.Clear();
            AnsiConsole.Write(new Rule("[yellow]Help: Health Check[/]").RuleStyle("grey").Centered());
            
            AnsiConsole.WriteLine();
            AnsiConsole.Write(new Panel(
                "[bold]Health Check[/]\n\n" +
                "This feature verifies that all required components of the InvoiceMailer application are properly configured.\n\n" +
                "[bold]What is checked:[/]\n" +
                "• Existence of required folders (invoices, logs)\n" +
                "• Existence and format of the recipients CSV file\n" +
                "• Proper configuration of email settings (TenantId, ClientId)\n" +
                "• Microsoft Graph API connectivity\n\n" +
                "[bold]Tips:[/]\n" +
                "• You can specify custom paths for the invoices folder and recipients file\n" +
                "• If any check fails, review the error message and fix the issue\n" +
                "• Use the Configure Settings option to update the application configuration\n\n" +
                "[bold]Security Note:[/]\n" +
                "• All credentials are loaded securely from appsettings.json\n" +
                "• Authentication uses interactive login for enhanced security\n" +
                "• No client secrets are stored in the application"
            ).Border(BoxBorder.Rounded).Padding(1, 1));
            
            Console.WriteLine("\nPress any key to return...");
            Console.ReadKey();
            AnsiConsole.Clear();
        }
        
        /// <summary>
        /// Perform health check with proper phase separation
        /// </summary>
        private static async Task PerformHealthCheck()
        {
            // PHASE 1: SETUP - Collect all necessary user input upfront
            AnsiConsole.Clear();
            AnsiConsole.Write(new Rule("[yellow]Health Check - Setup[/]").RuleStyle("grey").Centered());
            AnsiConsole.WriteLine();
            
            var invoicesFolder = AnsiConsole.Prompt(
                new TextPrompt<string>("Enter invoices folder path (or press Enter for default):")
                    .AllowEmpty());

            var recipientsFile = AnsiConsole.Prompt(
                new TextPrompt<string>("Enter recipients file path (or press Enter for default):")
                    .AllowEmpty());

            // Add confirmation with back option
            AnsiConsole.WriteLine();
            var proceedChoice = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("Ready to proceed?")
                    .PageSize(3)
                    .AddChoices(
                        "Proceed with Health Check",
                        "Back to Previous Screen"
                    ));
            
            if (proceedChoice == "Back to Previous Screen")
            {
                AnsiConsole.Clear();
                return;
            }

            // PHASE 2: PROCESSING - After all user input is collected
            AnsiConsole.Clear();
            AnsiConsole.Write(new Rule("[yellow]Health Check - Running[/]").RuleStyle("grey").Centered());
            AnsiConsole.WriteLine();
            
            bool success = false;
            await AnsiConsole.Status()
                .StartAsync("Running health check...", async ctx =>
                {
                    ctx.Spinner(Spinner.Known.Dots);
                    ctx.Status("Checking system...");

                    // Use the controller to run the health check
                    success = await _controller.RunHealthCheck(invoicesFolder, recipientsFile);
                });

            // PHASE 3: RESULTS - Display final results
            AnsiConsole.WriteLine();
            if (success)
            {
                AnsiConsole.MarkupLine("[green]Health check completed successfully![/]");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]Health check completed with issues. See log for details.[/]");
            }
            AnsiConsole.MarkupLine("[blue]Use the 'View Log' option to see detailed information.[/]");
            Console.WriteLine("Press any key to return to the previous screen...");
            Console.ReadKey();
            AnsiConsole.Clear();
        }

        private static async Task ConfigureSettings()
        {
            bool goBack = false;
            
            while (!goBack)
            {
                AnsiConsole.Clear();
                AnsiConsole.Write(new Rule("[yellow]Configuration Settings[/]").RuleStyle("grey").Centered());
                AnsiConsole.WriteLine();
                
                var choice = AnsiConsole.Prompt(
                    new SelectionPrompt<string>()
                        .Title("Configuration Options:")
                        .PageSize(10)
                        .AddChoices(
                            "Edit Settings",
                            "Help",
                            "Back to Main Menu"));
                
                switch (choice)
                {
                    case "Edit Settings":
                        EditSettings();
                        break;
                    case "Help":
                        ShowConfigurationHelp();
                        break;
                    case "Back to Main Menu":
                        goBack = true;
                        break;
                }
            }
            
            AnsiConsole.Clear();
            await Task.CompletedTask;
        }
        
        private static void ShowConfigurationHelp()
        {
            AnsiConsole.Clear();
            AnsiConsole.Write(new Rule("[yellow]Help: Configuration Settings[/]").RuleStyle("grey").Centered());
            
            AnsiConsole.WriteLine();
            AnsiConsole.Write(new Panel(
                "[bold]Configuration Settings[/]\n\n" +
                "This feature allows you to view and modify the application settings.\n\n" +
                "[bold]Available Settings:[/]\n" +
                "[underline]Authentication Settings:[/]\n" +
                "• [green]Test Mode[/]: When enabled, emails are logged but not actually sent\n" +
                "• [green]Azure AD Tenant ID[/]: Your Microsoft 365 tenant identifier\n" +
                "• [green]Azure AD Client ID[/]: The application ID registered in Azure AD\n" +
                "• [green]Default Sender Email[/]: Optional override for the sender email address\n\n" +
                "[underline]File Path Settings:[/]\n" +
                "• [green]Invoices Folder Path[/]: Where to look for invoice files\n" +
                "• [green]Recipients CSV File Path[/]: The CSV file containing recipient information\n\n" +
                "[underline]Scanner Settings:[/]\n" +
                "• [green]Invoice Key Pattern[/]: The regex pattern used to extract invoice keys from filenames\n" +
                "  (e.g., 'INV\\d+' extracts 'INV12345' from a filename like 'INV12345-March2023.pdf')\n\n" +
                "[bold]Authentication:[/]\n" +
                "• This application uses interactive authentication with Microsoft Identity\n" +
                "• Users will be prompted to sign in with their organizational account\n" +
                "• Emails will be sent on behalf of the authenticated user by default\n" +
                "• You can set a default sender override in the configuration\n" +
                "• You can also override the sender at runtime for each session\n\n" +
                "[bold]File Path Configuration:[/]\n" +
                "• Paths can be relative (e.g., 'invoices/') or absolute (e.g., '/Users/name/invoices/')\n" +
                "• The file browser allows you to easily navigate and select folders/files\n" +
                "• The application can create folders and files if they don't exist\n" +
                "• Recipients CSV file should have the format: 'InvoiceKey,Email'\n\n" +
                "[bold]Tips:[/]\n" +
                "• Keep Test Mode enabled during development and testing\n" +
                "• Obtain Azure AD credentials from your Microsoft 365 administrator\n" +
                "• Ensure your user account has proper permissions for sending emails\n" +
                "• Setting a default sender email is useful for shared mailbox scenarios\n\n" +
                "[bold]Security Note:[/]\n" +
                "• All credentials are loaded securely from appsettings.json\n" +
                "• User authentication is handled securely through Microsoft Identity\n" +
                "• NEVER share the appsettings.json file or commit it to public repositories"
            ).Border(BoxBorder.Rounded).Padding(1, 1));
            
            Console.WriteLine("\nPress any key to return...");
            Console.ReadKey();
            AnsiConsole.Clear();
        }
        
        private static void EditSettings()
        {
            // PHASE 1: PREPARATION - Check for appsettings.json
            AnsiConsole.Clear();
            AnsiConsole.Write(new Rule("[yellow]Edit Configuration[/]").RuleStyle("grey").Centered());
            AnsiConsole.WriteLine();
            
            // Get appsettings.json path
            var appSettingsPath = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");
            
            if (!File.Exists(appSettingsPath))
            {
                AnsiConsole.MarkupLine("[red]Error:[/] appsettings.json not found!");
                Console.WriteLine("Press any key to return to the previous screen...");
                Console.ReadKey();
                return;
            }

            // Load configuration
            var configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();

            // Display and allow editing of key settings
            AnsiConsole.Write(new Panel(
                "Edit your InvoiceMailer configuration settings.\n\n" +
                "[bold][yellow]INTERACTIVE AUTHENTICATION:[/][/]\n" +
                "This application uses interactive authentication with Microsoft Identity Platform.\n" +
                "• You will be prompted to sign in with your organization account when running the application\n" +
                "• Emails will be sent from your authenticated account by default\n" +
                "• You can set a default sender email override in the configuration\n" +
                "• You can also override the sender email at runtime if needed\n\n" +
                "[bold][yellow]FILE PATHS:[/][/]\n" +
                "• Configure paths for invoice files and recipient data\n" +
                "• Use the file browser to easily select folders and files\n" +
                "• Paths can be relative or absolute"
            ).Border(BoxBorder.Rounded)
            .Padding(1, 1));

            AnsiConsole.WriteLine();
            
            // PHASE 2: USER INPUT - Collect all user decisions and inputs
            
            // Select which settings to edit, with back option
            var settingGroup = AnsiConsole.Prompt(
                new SelectionPrompt<string>()
                    .Title("What type of settings would you like to edit?")
                    .PageSize(10)
                    .AddChoices(
                        "Authentication Settings",
                        "File Path Settings",
                        "Scanner Settings",
                        "Back to Previous Screen"
                    ));
            
            if (settingGroup == "Back to Previous Screen")
            {
                AnsiConsole.Clear();
                return;
            }
            
            bool savingRequired = false;
            string? invoicesFolderPath = null;
            string? recipientsFilePath = null;
            string? invoiceKeyPattern = null;
            bool createInvoicesFolder = false;
            bool createRecipientsParentDir = false;
            bool createRecipientsFile = false;
            string? recipientsParentDir = null;
            
            // Email and authentication settings
            bool testMode = configuration.GetValue<bool>("Email:TestMode");
            string tenantId = configuration.GetSection("Email").GetValue<string>("TenantId") ?? string.Empty;
            string clientId = configuration.GetSection("Email").GetValue<string>("ClientId") ?? string.Empty;
            string defaultSenderEmail = configuration.GetSection("Email").GetValue<string>("DefaultSenderEmail") ?? string.Empty;

            if (settingGroup == "Authentication Settings")
            {
                AnsiConsole.WriteLine();
                AnsiConsole.Write(new Rule("[yellow]Authentication Settings[/]").RuleStyle("grey").Centered());
                AnsiConsole.WriteLine();
                
                testMode = AnsiConsole.Confirm($"Test Mode (currently: {testMode})", testMode);
                
                tenantId = AnsiConsole.Prompt(
                    new TextPrompt<string>($"Azure AD Tenant ID:")
                        .DefaultValue(tenantId)
                        .PromptStyle("green"));

                clientId = AnsiConsole.Prompt(
                    new TextPrompt<string>($"Azure AD Client ID:")
                        .DefaultValue(clientId)
                        .PromptStyle("green"));
                
                var setSenderEmail = AnsiConsole.Confirm("Do you want to set a default sender email?", !string.IsNullOrEmpty(defaultSenderEmail));
                
                if (setSenderEmail)
                {
                    defaultSenderEmail = AnsiConsole.Prompt(
                        new TextPrompt<string>("Default Sender Email:")
                            .DefaultValue(defaultSenderEmail)
                            .Validate(email => 
                            {
                                if (string.IsNullOrWhiteSpace(email))
                                    return ValidationResult.Error("Email cannot be empty");
                                if (!email.Contains('@'))
                                    return ValidationResult.Error("Invalid email format");
                                return ValidationResult.Success();
                            })
                            .PromptStyle("green"));
                }
                else
                {
                    defaultSenderEmail = string.Empty;
                }
                
                savingRequired = true;
            }
            else if (settingGroup == "File Path Settings")
            {
                AnsiConsole.WriteLine();
                AnsiConsole.Write(new Rule("[yellow]File Path Settings[/]").RuleStyle("grey").Centered());
                AnsiConsole.WriteLine();
                
                // Get current paths
                string currentInvoicesPath = _controller.GetInvoicesFolderPath();
                string currentRecipientsPath = _controller.GetRecipientsFilePath();
                
                // Show the current paths
                AnsiConsole.MarkupLine($"Current Invoices Folder: [blue]{currentInvoicesPath}[/]");
                AnsiConsole.MarkupLine($"Current Recipients File: [green]{currentRecipientsPath}[/]");
                AnsiConsole.WriteLine();
                
                // Option to edit invoices folder
                if (AnsiConsole.Confirm("Would you like to change the invoices folder?", false))
                {
                    var pathMode = AnsiConsole.Prompt(
                        new SelectionPrompt<string>()
                            .Title("How would you like to set the path?")
                            .PageSize(4)
                            .AddChoices(
                                "Browse for folder",
                                "Enter path manually",
                                "Reset to default",
                                "Cancel"
                            ));
                    
                    if (pathMode == "Cancel")
                    {
                        // Skip this setting and continue
                    }
                    else if (pathMode == "Browse for folder")
                    {
                            var selectedFolder = NativeFileBrowser.BrowseForFolder(
                                "Select Invoices Folder", 
                                currentInvoicesPath);
                            
                            if (selectedFolder != null)
                            {
                                invoicesFolderPath = selectedFolder;
                                AnsiConsole.MarkupLine($"[green]New invoices folder set to: {invoicesFolderPath}[/]");
                                savingRequired = true;
                            
                            // Check if directory exists and get confirmation if needed
                            if (!Directory.Exists(invoicesFolderPath))
                            {
                                createInvoicesFolder = AnsiConsole.Confirm($"Directory '{invoicesFolderPath}' does not exist. Create it?", true);
                            }
                        }
                    }
                    else if (pathMode == "Enter path manually")
                    {
                            var enteredPath = AnsiConsole.Prompt(
                                new TextPrompt<string>("Enter the invoices folder path:")
                                    .DefaultValue(currentInvoicesPath)
                                    .Validate(path =>
                                    {
                                        if (string.IsNullOrWhiteSpace(path))
                                            return ValidationResult.Error("Path cannot be empty");
                                        return ValidationResult.Success();
                                    }));
                            
                            invoicesFolderPath = enteredPath;
                            savingRequired = true;
                            
                        // Check if directory exists and get confirmation if needed
                            if (!Directory.Exists(invoicesFolderPath))
                            {
                            createInvoicesFolder = AnsiConsole.Confirm($"Directory '{invoicesFolderPath}' does not exist. Create it?", true);
                        }
                    }
                    else if (pathMode == "Reset to default")
                    {
                            invoicesFolderPath = "invoices";
                            AnsiConsole.MarkupLine("[green]Reset invoices folder to default: 'invoices'[/]");
                            savingRequired = true;
                        
                        // Check if default directory exists and get confirmation if needed
                        if (!Directory.Exists(invoicesFolderPath))
                        {
                            createInvoicesFolder = AnsiConsole.Confirm($"Default directory '{invoicesFolderPath}' does not exist. Create it?", true);
                        }
                    }
                }
                
                AnsiConsole.WriteLine();
                
                // Option to edit recipients file
                if (AnsiConsole.Confirm("Would you like to change the recipients file?", false))
                {
                    var pathMode = AnsiConsole.Prompt(
                        new SelectionPrompt<string>()
                            .Title("How would you like to set the path?")
                            .PageSize(4)
                            .AddChoices(
                                "Browse for file",
                                "Enter path manually",
                                "Reset to default",
                                "Cancel"
                            ));
                    
                    if (pathMode == "Cancel")
                    {
                        // Skip this setting and continue
                    }
                    else if (pathMode == "Browse for file")
                    {
                            var selectedFile = NativeFileBrowser.BrowseForFile(
                                "Select Recipients CSV File",
                                "CSV Files (*.csv)|*.csv", 
                                Path.GetDirectoryName(currentRecipientsPath));
                            
                            if (selectedFile != null)
                            {
                                recipientsFilePath = selectedFile;
                                AnsiConsole.MarkupLine($"[green]New recipients file set to: {recipientsFilePath}[/]");
                                savingRequired = true;
                            
                            // Check if file exists and get confirmation if needed
                            if (!File.Exists(recipientsFilePath))
                            {
                                createRecipientsFile = AnsiConsole.Confirm($"File '{recipientsFilePath}' does not exist. Create an empty CSV file?", true);
                                
                                // Check if parent directory exists
                                recipientsParentDir = Path.GetDirectoryName(recipientsFilePath);
                                if (!string.IsNullOrEmpty(recipientsParentDir) && !Directory.Exists(recipientsParentDir))
                                {
                                    createRecipientsParentDir = AnsiConsole.Confirm($"Directory '{recipientsParentDir}' does not exist. Create it?", true);
                                }
                            }
                        }
                    }
                    else if (pathMode == "Enter path manually")
                    {
                            var enteredPath = AnsiConsole.Prompt(
                                new TextPrompt<string>("Enter the recipients file path:")
                                    .DefaultValue(currentRecipientsPath)
                                    .Validate(path =>
                                    {
                                        if (string.IsNullOrWhiteSpace(path))
                                            return ValidationResult.Error("Path cannot be empty");
                                        return ValidationResult.Success();
                                    }));
                            
                            recipientsFilePath = enteredPath;
                            savingRequired = true;
                            
                        // Check parent directory and file (don't create yet, just mark for creation later)
                        recipientsParentDir = Path.GetDirectoryName(recipientsFilePath);
                        if (!string.IsNullOrEmpty(recipientsParentDir) && !Directory.Exists(recipientsParentDir))
                        {
                            createRecipientsParentDir = AnsiConsole.Confirm($"Directory '{recipientsParentDir}' does not exist. Create it?", true);
                        }
                        
                        if (!File.Exists(recipientsFilePath))
                        {
                            createRecipientsFile = AnsiConsole.Confirm($"File '{recipientsFilePath}' does not exist. Create an empty CSV file?", true);
                        }
                    }
                    else if (pathMode == "Reset to default")
                    {
                        recipientsFilePath = "recipients.csv";
                        AnsiConsole.MarkupLine("[green]Reset recipients file to default: 'recipients.csv'[/]");
                        savingRequired = true;
                        
                        // Check if default file exists
                        if (!File.Exists(recipientsFilePath))
                        {
                            createRecipientsFile = AnsiConsole.Confirm($"Default file '{recipientsFilePath}' does not exist. Create an empty CSV file?", true);
                        }
                    }
                }
            }
            else if (settingGroup == "Scanner Settings")
            {
                AnsiConsole.WriteLine();
                AnsiConsole.Write(new Rule("[yellow]Scanner Settings[/]").RuleStyle("grey").Centered());
                AnsiConsole.WriteLine();
                
                // Add information panel about invoice key patterns
                AnsiConsole.Write(new Panel(
                    "[bold]Invoice Key Pattern[/]\n\n" +
                    "This setting controls the regular expression pattern used to extract invoice keys from filenames.\n\n" +
                    "[green]Examples:[/]\n" +
                    "• [yellow]INV\\d+[/] - Matches 'INV' followed by one or more digits (e.g., INV12345)\n" +
                    "• [yellow]INVOICE_\\d{4}[/] - Matches 'INVOICE_' followed by exactly 4 digits\n" +
                    "• [yellow]\\d{5}[/] - Matches any 5-digit number\n" +
                    "• [yellow]PR-\\d+-\\d+[/] - Matches patterns like 'PR-123-456'\n\n" +
                    "The matched text will be used to look up recipient email addresses in the CSV file."
                ).Border(BoxBorder.Rounded).Padding(1, 1));
                
                AnsiConsole.WriteLine();
                
                // Get current scanner settings
                string currentInvoiceKeyPattern = _controller.GetInvoiceKeyPattern();
                
                // Show the current pattern
                AnsiConsole.MarkupLine($"Current Invoice Key Pattern: [blue]{currentInvoiceKeyPattern}[/]");
                AnsiConsole.WriteLine();
                
                // Option to edit invoice key pattern
                if (AnsiConsole.Confirm("Would you like to change the invoice key pattern?", false))
                {
                    var newPattern = AnsiConsole.Prompt(
                        new TextPrompt<string>("Enter the new invoice key pattern:")
                            .DefaultValue(currentInvoiceKeyPattern)
                            .Validate(pattern =>
                            {
                                if (string.IsNullOrWhiteSpace(pattern))
                                    return ValidationResult.Error("Pattern cannot be empty");
                                
                                // Test if it's a valid regex pattern
                                try 
                                {
                                    var _ = new System.Text.RegularExpressions.Regex(pattern);
                                    return ValidationResult.Success();
                                }
                                catch (Exception)
                                {
                                    return ValidationResult.Error("Invalid regular expression pattern");
                                }
                            }));
                    
                    invoiceKeyPattern = newPattern;
                    savingRequired = true;
                    
                    AnsiConsole.MarkupLine($"[green]New invoice key pattern set to: {invoiceKeyPattern}[/]");
                }
            }

            // Confirm saving changes
            bool saveConfirmed = false;
            if (savingRequired)
            {
                AnsiConsole.WriteLine();
                
                var saveChoice = AnsiConsole.Prompt(
                    new SelectionPrompt<string>()
                        .Title("Would you like to save these changes?")
                        .PageSize(3)
                        .AddChoices(
                            "Save Changes",
                            "Discard Changes",
                            "Back to Previous Screen"
                        ));
                
                if (saveChoice == "Back to Previous Screen")
                {
                    AnsiConsole.Clear();
                    return;
                }
                
                saveConfirmed = (saveChoice == "Save Changes");
            }
            
            // PHASE 3: PROCESSING - Perform file operations and saving inside Status spinner
            bool success = false;
            
            if (savingRequired && saveConfirmed)
            {
                AnsiConsole.Status()
                    .Start("Processing configuration changes...", ctx => 
                    {
                        ctx.Spinner(Spinner.Known.Dots);
                        
                        // Create directories if needed
                        if (createInvoicesFolder && invoicesFolderPath != null)
                        {
                            ctx.Status($"Creating directory: {invoicesFolderPath}");
                            try
                            {
                                Directory.CreateDirectory(invoicesFolderPath);
                                _logEntries.Add(($"Created directory: {invoicesFolderPath}", InvoiceMailerController.LogLevel.Success));
                                    }
                                    catch (Exception ex)
                                    {
                                _logEntries.Add(($"Error creating directory: {ex.Message}", InvoiceMailerController.LogLevel.Error));
                            }
                        }
                        
                        if (createRecipientsParentDir && recipientsParentDir != null)
                        {
                            ctx.Status($"Creating directory: {recipientsParentDir}");
                            try
                            {
                                Directory.CreateDirectory(recipientsParentDir);
                                _logEntries.Add(($"Created directory: {recipientsParentDir}", InvoiceMailerController.LogLevel.Success));
                            }
                            catch (Exception ex)
                            {
                                _logEntries.Add(($"Error creating directory: {ex.Message}", InvoiceMailerController.LogLevel.Error));
                            }
                        }
                        
                        if (createRecipientsFile && recipientsFilePath != null)
                        {
                            ctx.Status($"Creating file: {recipientsFilePath}");
                            try
                            {
                                File.WriteAllText(recipientsFilePath, "InvoiceKey,Email\n");
                                _logEntries.Add(($"Created empty CSV file with header: {recipientsFilePath}", InvoiceMailerController.LogLevel.Success));
                            }
                            catch (Exception ex)
                            {
                                _logEntries.Add(($"Error creating file: {ex.Message}", InvoiceMailerController.LogLevel.Error));
                            }
                        }
                        
                        // Save configuration
                        ctx.Status("Saving configuration to appsettings.json");
                        success = _controller.SaveConfiguration(
                        testMode, 
                        tenantId, 
                        clientId, 
                        defaultSenderEmail,
                        invoicesFolderPath,
                            recipientsFilePath,
                            invoiceKeyPattern);
                    
                    if (success)
                    {
                            _logEntries.Add(("Settings saved successfully!", InvoiceMailerController.LogLevel.Success));
                    }
                    else
                    {
                            _logEntries.Add(("Failed to save settings.", InvoiceMailerController.LogLevel.Error));
                        }
                    });
                
                // PHASE 4: RESULTS
                AnsiConsole.WriteLine();
                
                // Display the latest log entries
                for (int i = Math.Max(0, _logEntries.Count - 5); i < _logEntries.Count; i++)
                {
                    var entry = _logEntries[i];
                    string color = entry.Level switch
                    {
                        InvoiceMailerController.LogLevel.Success => "green",
                        InvoiceMailerController.LogLevel.Warning => "yellow",
                        InvoiceMailerController.LogLevel.Error => "red",
                        _ => "blue"
                    };
                    AnsiConsole.MarkupLine($"[{color}]{entry.Message}[/]");
                }
            }
            else if (savingRequired)
                {
                    AnsiConsole.MarkupLine("[yellow]Settings were not saved.[/]");
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]No changes were made to the settings.[/]");
            }

            Console.WriteLine("Press any key to return to the previous screen...");
            Console.ReadKey();
            AnsiConsole.Clear();
        }
    }
}
