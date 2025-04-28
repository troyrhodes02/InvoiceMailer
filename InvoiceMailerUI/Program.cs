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
                "• ClientSecret must be rotated annually per IT security policy"
            ).Border(BoxBorder.Rounded).Padding(1, 1));
            
            Console.WriteLine("\nPress any key to return...");
            Console.ReadKey();
            AnsiConsole.Clear();
        }
        
        private static async Task ProcessInvoices(bool dryRun)
        {
            var invoicesFolder = AnsiConsole.Prompt(
                new TextPrompt<string>("Enter invoices folder path (or press Enter for default):")
                    .AllowEmpty());

            var recipientsFile = AnsiConsole.Prompt(
                new TextPrompt<string>("Enter recipients file path (or press Enter for default):")
                    .AllowEmpty());

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

            bool success = false;
            await AnsiConsole.Status()
                .StartAsync("Processing invoices...", async ctx =>
                {
                    ctx.Spinner(Spinner.Known.Dots);
                    ctx.Status("Running InvoiceMailer...");

                    // Use the controller to run the appropriate process
                    if (dryRun)
                    {
                        success = await _controller.RunDryRun(invoicesFolder, recipientsFile);
                    }
                    else
                    {
                        success = await _controller.RunRealSend(invoicesFolder, recipientsFile);
                    }
                });

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
            Console.WriteLine("Press any key to continue...");
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
                "• Proper configuration of email settings (TenantId, ClientId, etc.)\n\n" +
                "[bold]Tips:[/]\n" +
                "• You can specify custom paths for the invoices folder and recipients file\n" +
                "• If any check fails, review the error message and fix the issue\n" +
                "• Use the Configure Settings option to update the application configuration\n\n" +
                "[bold]Security Note:[/]\n" +
                "• All credentials are loaded securely from appsettings.json\n" +
                "• ClientSecret must be rotated annually per IT security policy"
            ).Border(BoxBorder.Rounded).Padding(1, 1));
            
            Console.WriteLine("\nPress any key to return...");
            Console.ReadKey();
            AnsiConsole.Clear();
        }
        
        private static async Task PerformHealthCheck()
        {
            var invoicesFolder = AnsiConsole.Prompt(
                new TextPrompt<string>("Enter invoices folder path (or press Enter for default):")
                    .AllowEmpty());

            var recipientsFile = AnsiConsole.Prompt(
                new TextPrompt<string>("Enter recipients file path (or press Enter for default):")
                    .AllowEmpty());

            bool success = false;
            await AnsiConsole.Status()
                .StartAsync("Running health check...", async ctx =>
                {
                    ctx.Spinner(Spinner.Known.Dots);
                    ctx.Status("Checking system...");

                    // Use the controller to run the health check
                    success = await _controller.RunHealthCheck(invoicesFolder, recipientsFile);
                });

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
            Console.WriteLine("Press any key to continue...");
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
                "• [green]Test Mode[/]: When enabled, emails are logged but not actually sent\n" +
                "• [green]Azure AD Tenant ID[/]: Your Microsoft 365 tenant identifier\n" +
                "• [green]Azure AD Client ID[/]: The application ID registered in Azure AD\n" +
                "• [green]Azure AD Client Secret[/]: The secret key for authentication\n" +
                "• [green]Sender Email[/]: The email address used as the sender\n\n" +
                "[bold]Tips:[/]\n" +
                "• Keep Test Mode enabled during development and testing\n" +
                "• Obtain Azure AD credentials from your Microsoft 365 administrator\n" +
                "• Ensure the sender email has proper permissions in your Microsoft 365 tenant\n\n" +
                "[bold]Security Note:[/]\n" +
                "• All credentials are loaded securely from appsettings.json\n" +
                "• ClientSecret must be rotated annually per IT security policy\n" +
                "• NEVER share the appsettings.json file or commit it to public repositories"
            ).Border(BoxBorder.Rounded).Padding(1, 1));
            
            Console.WriteLine("\nPress any key to return...");
            Console.ReadKey();
            AnsiConsole.Clear();
        }
        
        private static void EditSettings()
        {
            // Get appsettings.json path
            var appSettingsPath = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");
            
            if (!File.Exists(appSettingsPath))
            {
                AnsiConsole.MarkupLine("[red]Error:[/] appsettings.json not found!");
                Console.WriteLine("Press any key to continue...");
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
                "This will display the current settings from appsettings.json.\n" +
                "You can edit the settings and save them back to the file.\n\n" +
                "[bold][yellow]SECURITY WARNING:[/][/] Never share these credentials or commit them to public repositories."
            ).Border(BoxBorder.Rounded)
            .Padding(1, 1));

            // Email settings
            var emailConfig = configuration.GetSection("Email");
            var testMode = configuration.GetValue<bool>("Email:TestMode");
            var tenantId = emailConfig.GetValue<string>("TenantId") ?? string.Empty;
            var clientId = emailConfig.GetValue<string>("ClientId") ?? string.Empty;
            var clientSecret = emailConfig.GetValue<string>("ClientSecret") ?? string.Empty;
            var senderEmail = emailConfig.GetValue<string>("SenderEmail") ?? string.Empty;

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

            clientSecret = AnsiConsole.Prompt(
                new TextPrompt<string>($"Azure AD Client Secret:")
                    .DefaultValue(clientSecret)
                    .Secret()
                    .PromptStyle("green"));

            senderEmail = AnsiConsole.Prompt(
                new TextPrompt<string>($"Sender Email Address:")
                    .DefaultValue(senderEmail)
                    .PromptStyle("green"));

            AnsiConsole.WriteLine();
            var saveConfirmed = AnsiConsole.Confirm("Save these settings to appsettings.json?", true);
            
            if (saveConfirmed)
            {
                bool success = _controller.SaveConfiguration(testMode, tenantId, clientId, clientSecret, senderEmail);
                
                if (success)
                {
                    AnsiConsole.MarkupLine("[green]Settings saved successfully![/]");
                    AnsiConsole.MarkupLine("[yellow]Remember that ClientSecret must be rotated annually per IT security policy.[/]");
                }
                else
                {
                    AnsiConsole.MarkupLine("[red]Failed to save settings. See log for details.[/]");
                }
            }
            else
            {
                AnsiConsole.MarkupLine("[yellow]Settings were not saved.[/]");
            }

            Console.WriteLine("Press any key to continue...");
            Console.ReadKey();
            AnsiConsole.Clear();
        }
    }
}
