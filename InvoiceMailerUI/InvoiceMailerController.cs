using System;
using System.Collections.Generic;
using System.CommandLine;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Spectre.Console;
using System.Text.Json;
using System.Text.Json.Nodes;

namespace InvoiceMailerUI
{
    /// <summary>
    /// Controller for InvoiceMailer operations, separating business logic from UI
    /// </summary>
    public class InvoiceMailerController
    {
        private readonly IConfiguration _configuration;
        private Action<string, LogLevel>? _logCallback;

        public enum LogLevel
        {
            Info,
            Success,
            Warning,
            Error
        }

        public InvoiceMailerController()
        {
            // Load configuration
            _configuration = new ConfigurationBuilder()
                .SetBasePath(Directory.GetCurrentDirectory())
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
                .Build();
        }

        /// <summary>
        /// Set a callback for real-time logging
        /// </summary>
        public void SetLogCallback(Action<string, LogLevel> logCallback)
        {
            _logCallback = logCallback;
        }

        /// <summary>
        /// Log a message using the callback if available
        /// </summary>
        private void Log(string message, LogLevel level = LogLevel.Info)
        {
            _logCallback?.Invoke(message, level);
        }

        /// <summary>
        /// Check if setup wizard is needed and run it if required
        /// </summary>
        public async Task<bool> RunSetupWizard()
        {
            Log("Checking if setup wizard is needed...", LogLevel.Info);
            
            bool setupRequired = false;
            string appSettingsPath = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");
            
            // Check if appsettings.json exists
            if (!File.Exists(appSettingsPath))
            {
                Log("appsettings.json not found. Setup wizard will create it.", LogLevel.Warning);
                setupRequired = true;
            }
            else
            {
                try
                {
                    // Check if required settings are present
                    var emailConfig = _configuration.GetSection("Email");
                    var configTenantId = emailConfig.GetValue<string>("TenantId") ?? string.Empty;
                    var configClientId = emailConfig.GetValue<string>("ClientId") ?? string.Empty;
                    var configClientSecret = emailConfig.GetValue<string>("ClientSecret") ?? string.Empty;
                    
                    if (string.IsNullOrWhiteSpace(configTenantId) || 
                        string.IsNullOrWhiteSpace(configClientId) || 
                        string.IsNullOrWhiteSpace(configClientSecret))
                    {
                        Log("Critical email settings are missing. Setup wizard will configure them.", LogLevel.Warning);
                        setupRequired = true;
                    }
                }
                catch (Exception ex)
                {
                    Log($"Error reading appsettings.json: {ex.Message}", LogLevel.Error);
                    setupRequired = true;
                }
            }
            
            if (!setupRequired)
            {
                Log("Setup wizard not needed, configuration is complete.", LogLevel.Success);
                return true; // No setup needed, continue with normal operation
            }
            
            // Setup wizard is needed - gather information
            AnsiConsole.Clear();
            AnsiConsole.Write(new FigletText("Setup Wizard").Color(Color.Orange1));
            AnsiConsole.WriteLine();
            
            var panel = new Panel(
                "This wizard will help you configure critical settings for InvoiceMailer.\n\n" +
                "[bold][yellow]⚠️ SECURITY WARNING:[/][/] Please treat TenantId, ClientId, and ClientSecret as sensitive credentials.\n" +
                "Do not share these credentials outside of authorized personnel."
            ).Border(BoxBorder.Rounded)
            .Padding(1, 1);
            
            AnsiConsole.Write(panel);
            AnsiConsole.WriteLine();
            
            // Prompt for settings
            string wizardTenantId = AnsiConsole.Prompt(
                new TextPrompt<string>("Enter your [green]Azure AD Tenant ID[/]:")
                    .PromptStyle("green")
                    .ValidationErrorMessage("[red]The Tenant ID cannot be empty[/]")
                    .Validate(id => string.IsNullOrWhiteSpace(id) 
                        ? ValidationResult.Error("[red]The Tenant ID cannot be empty[/]") 
                        : ValidationResult.Success()));
            
            string wizardClientId = AnsiConsole.Prompt(
                new TextPrompt<string>("Enter your [green]Azure AD Client ID[/]:")
                    .PromptStyle("green")
                    .ValidationErrorMessage("[red]The Client ID cannot be empty[/]")
                    .Validate(id => string.IsNullOrWhiteSpace(id) 
                        ? ValidationResult.Error("[red]The Client ID cannot be empty[/]") 
                        : ValidationResult.Success()));
            
            string wizardClientSecret = AnsiConsole.Prompt(
                new TextPrompt<string>("Enter your [green]Azure AD Client Secret[/]:")
                    .PromptStyle("green")
                    .Secret()
                    .ValidationErrorMessage("[red]The Client Secret cannot be empty[/]")
                    .Validate(secret => string.IsNullOrWhiteSpace(secret) 
                        ? ValidationResult.Error("[red]The Client Secret cannot be empty[/]") 
                        : ValidationResult.Success()));
            
            string senderEmail = AnsiConsole.Prompt(
                new TextPrompt<string>("Enter your [green]Sender Email Address[/] (or press Enter for default):")
                    .PromptStyle("green")
                    .DefaultValue("default@example.com")
                    .AllowEmpty());
            
            AnsiConsole.WriteLine();
            bool success = await SaveSetupConfiguration(wizardTenantId, wizardClientId, wizardClientSecret, senderEmail);
            
            if (success)
            {
                AnsiConsole.MarkupLine("[green]Setup completed successfully![/]");
                AnsiConsole.MarkupLine("[blue]Configuration saved to appsettings.json[/]");
                Log("Setup wizard completed successfully", LogLevel.Success);
                
                // Wait for user to acknowledge completion
                AnsiConsole.WriteLine();
                Console.WriteLine("Press any key to continue to the application...");
                Console.ReadKey();
                AnsiConsole.Clear();
                return true;
            }
            else
            {
                AnsiConsole.MarkupLine("[red]Failed to save configuration settings.[/]");
                Log("Setup wizard failed to save configuration", LogLevel.Error);
                
                // Wait for user to acknowledge error
                AnsiConsole.WriteLine();
                Console.WriteLine("Press any key to exit the application...");
                Console.ReadKey();
                return false;
            }
        }
        
        /// <summary>
        /// Save configuration settings from the setup wizard
        /// </summary>
        private async Task<bool> SaveSetupConfiguration(string wizardTenantId, string wizardClientId, string wizardClientSecret, string senderEmail)
        {
            Log("Saving configuration from setup wizard...", LogLevel.Info);
            
            try
            {
                string appSettingsPath = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");
                
                // Load existing config if it exists, or create a new one
                JsonNode configRoot;
                if (File.Exists(appSettingsPath))
                {
                    string existingJson = await File.ReadAllTextAsync(appSettingsPath);
                    configRoot = JsonNode.Parse(existingJson) ?? new JsonObject();
                }
                else
                {
                    configRoot = new JsonObject
                    {
                        ["ApplicationSettings"] = new JsonObject
                        {
                            ["ApplicationName"] = "InvoiceMailer",
                            ["Environment"] = "Development"
                        },
                        ["InvoiceScanner"] = new JsonObject
                        {
                            ["DefaultPattern"] = "INV\\d+",
                            ["ScanPath"] = "invoices",
                            ["CaseInsensitive"] = true
                        },
                        ["RecipientLookup"] = new JsonObject
                        {
                            ["CsvPath"] = "recipients.csv"
                        },
                        ["Serilog"] = new JsonObject
                        {
                            ["Using"] = new JsonArray { "Serilog.Sinks.Console", "Serilog.Sinks.File" },
                            ["MinimumLevel"] = new JsonObject
                            {
                                ["Default"] = "Information",
                                ["Override"] = new JsonObject
                                {
                                    ["Microsoft"] = "Warning",
                                    ["System"] = "Warning"
                                }
                            },
                            ["WriteTo"] = new JsonArray
                            {
                                new JsonObject { ["Name"] = "Console" },
                                new JsonObject
                                {
                                    ["Name"] = "File",
                                    ["Args"] = new JsonObject
                                    {
                                        ["path"] = "logs/log.txt",
                                        ["rollingInterval"] = "Day",
                                        ["retainedFileCountLimit"] = 7,
                                        ["outputTemplate"] = "{Timestamp:yyyy-MM-dd HH:mm:ss.fff zzz} [{Level:u3}] {Message:lj}{NewLine}{Exception}"
                                    }
                                }
                            },
                            ["Enrich"] = new JsonArray { "FromLogContext", "WithMachineName", "WithThreadId" }
                        }
                    };
                }
                
                // Create or update Email section
                if (configRoot["Email"] is null)
                {
                    configRoot["Email"] = new JsonObject();
                }
                
                var emailSection = configRoot["Email"].AsObject();
                emailSection["TenantId"] = wizardTenantId;
                emailSection["ClientId"] = wizardClientId;
                emailSection["ClientSecret"] = wizardClientSecret;
                emailSection["SenderEmail"] = senderEmail;
                
                // Set default values if not present
                if (emailSection["TestMode"] is null)
                {
                    emailSection["TestMode"] = false;
                }
                if (emailSection["DryRun"] is null)
                {
                    emailSection["DryRun"] = false;
                }
                
                // Write updated configuration to file
                var options = new JsonSerializerOptions { WriteIndented = true };
                await File.WriteAllTextAsync(appSettingsPath, configRoot.ToJsonString(options));
                
                Log("Configuration saved successfully from setup wizard", LogLevel.Success);
                return true;
            }
            catch (Exception ex)
            {
                Log($"Error saving configuration from setup wizard: {ex.Message}", LogLevel.Error);
                return false;
            }
        }

        /// <summary>
        /// Run InvoiceMailer in dry-run mode
        /// </summary>
        public async Task<bool> RunDryRun(string? invoicesFolder = null, string? recipientsFile = null)
        {
            Log("Starting dry-run mode...", LogLevel.Info);
            
            try
            {
                bool result = await RunInvoiceMailer(true, invoicesFolder, recipientsFile);
                if (result)
                {
                    Log("Dry-run completed successfully.", LogLevel.Success);
                }
                else
                {
                    Log("Dry-run completed with issues. Check the logs for details.", LogLevel.Warning);
                }
                return result;
            }
            catch (Exception ex)
            {
                Log($"Error during dry-run: {ex.Message}", LogLevel.Error);
                return false;
            }
        }

        /// <summary>
        /// Run InvoiceMailer in real-send mode
        /// </summary>
        public async Task<bool> RunRealSend(string? invoicesFolder = null, string? recipientsFile = null)
        {
            Log("Starting real-send mode...", LogLevel.Info);
            Log("WARNING: This will send actual emails to recipients!", LogLevel.Warning);
            
            try
            {
                bool result = await RunInvoiceMailer(false, invoicesFolder, recipientsFile);
                if (result)
                {
                    Log("Email sending completed successfully.", LogLevel.Success);
                }
                else
                {
                    Log("Email sending completed with issues. Check the logs for details.", LogLevel.Warning);
                }
                return result;
            }
            catch (Exception ex)
            {
                Log($"Error during email sending: {ex.Message}", LogLevel.Error);
                return false;
            }
        }

        /// <summary>
        /// Run InvoiceMailer's health check
        /// </summary>
        public async Task<bool> RunHealthCheck(string? invoicesFolder = null, string? recipientsFile = null)
        {
            Log("Starting health check...", LogLevel.Info);
            
            try
            {
                // Prepare arguments for health check
                var args = new List<string> { "--check" };

                if (!string.IsNullOrEmpty(invoicesFolder))
                {
                    args.Add("--invoices");
                    args.Add(invoicesFolder);
                    Log($"Using custom invoices folder: {invoicesFolder}", LogLevel.Info);
                }

                if (!string.IsNullOrEmpty(recipientsFile))
                {
                    args.Add("--recipients");
                    args.Add(recipientsFile);
                    Log($"Using custom recipients file: {recipientsFile}", LogLevel.Info);
                }

                // Run the health check
                var rootCommand = BuildRootCommand();
                int exitCode = await rootCommand.InvokeAsync(args.ToArray());
                
                if (exitCode == 0)
                {
                    Log("Health check completed successfully.", LogLevel.Success);
                    return true;
                }
                else
                {
                    Log($"Health check failed with exit code: {exitCode}", LogLevel.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log($"Error during health check: {ex.Message}", LogLevel.Error);
                return false;
            }
        }

        /// <summary>
        /// Verify and secure configuration settings
        /// </summary>
        public bool VerifyConfiguration()
        {
            Log("Verifying configuration settings...", LogLevel.Info);
            
            bool isValid = true;
            
            // Check if all required settings are present
            var emailConfig = _configuration.GetSection("Email");
            var tenantId = emailConfig.GetValue<string>("TenantId") ?? string.Empty;
            var clientId = emailConfig.GetValue<string>("ClientId") ?? string.Empty;
            var clientSecret = emailConfig.GetValue<string>("ClientSecret") ?? string.Empty;
            var senderEmail = emailConfig.GetValue<string>("SenderEmail") ?? string.Empty;

            if (string.IsNullOrWhiteSpace(tenantId))
            {
                Log("TenantId is not configured in appsettings.json", LogLevel.Error);
                isValid = false;
            }
            
            if (string.IsNullOrWhiteSpace(clientId))
            {
                Log("ClientId is not configured in appsettings.json", LogLevel.Error);
                isValid = false;
            }
            
            if (string.IsNullOrWhiteSpace(clientSecret))
            {
                Log("ClientSecret is not configured in appsettings.json", LogLevel.Error);
                isValid = false;
            }
            
            if (string.IsNullOrWhiteSpace(senderEmail))
            {
                Log("SenderEmail is not configured in appsettings.json", LogLevel.Error);
                isValid = false;
            }

            // Check for required folders
            var scannerConfig = _configuration.GetSection("InvoiceScanner");
            var configScanPath = scannerConfig.GetValue<string>("ScanPath") ?? "invoices";
            if (!Directory.Exists(configScanPath))
            {
                Log($"Creating invoices folder: {configScanPath}", LogLevel.Info);
                try
                {
                    Directory.CreateDirectory(configScanPath);
                }
                catch (Exception ex)
                {
                    Log($"Error creating invoices folder: {ex.Message}", LogLevel.Error);
                    isValid = false;
                }
            }
            
            // Check for logs folder
            var logsFolder = Path.Combine(Directory.GetCurrentDirectory(), "logs");
            if (!Directory.Exists(logsFolder))
            {
                Log($"Creating logs folder: {logsFolder}", LogLevel.Info);
                try
                {
                    Directory.CreateDirectory(logsFolder);
                }
                catch (Exception ex)
                {
                    Log($"Error creating logs folder: {ex.Message}", LogLevel.Error);
                    isValid = false;
                }
            }
            
            if (isValid)
            {
                Log("Configuration verification completed successfully.", LogLevel.Success);
            }
            else
            {
                Log("Configuration verification failed. Please check the errors above.", LogLevel.Error);
            }
            
            return isValid;
        }

        /// <summary>
        /// Save configuration settings
        /// </summary>
        public bool SaveConfiguration(bool testMode, string tenantId, string clientId, string clientSecret, string senderEmail)
        {
            Log("Saving configuration settings...", LogLevel.Info);
            
            try
            {
                var appSettingsPath = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");
                if (!File.Exists(appSettingsPath))
                {
                    Log("appsettings.json not found!", LogLevel.Error);
                    return false;
                }

                // Read the existing file content
                string json = File.ReadAllText(appSettingsPath);
                
                // Use json manipulation library or for simplicity here, just do some basic string replacements
                // Note: This is not a robust way to modify JSON, but works for demo purposes
                
                // Update settings
                var settings = new Dictionary<string, string>
                {
                    { "\"TestMode\": true", $"\"TestMode\": {testMode.ToString().ToLower()}" },
                    { "\"TestMode\": false", $"\"TestMode\": {testMode.ToString().ToLower()}" },
                    { "\"TenantId\": \"\"", $"\"TenantId\": \"{tenantId}\"" },
                    { "\"ClientId\": \"\"", $"\"ClientId\": \"{clientId}\"" },
                    { "\"ClientSecret\": \"\"", $"\"ClientSecret\": \"{clientSecret}\"" },
                    { "\"SenderEmail\": \"\"", $"\"SenderEmail\": \"{senderEmail}\"" }
                };

                foreach (var setting in settings)
                {
                    json = json.Replace(setting.Key, setting.Value);
                }
                
                // Write the updated content back to the file
                File.WriteAllText(appSettingsPath, json);
                
                Log("Configuration saved successfully.", LogLevel.Success);
                Log("NOTE: ClientSecret must be rotated annually per IT security policy.", LogLevel.Warning);
                return true;
            }
            catch (Exception ex)
            {
                Log($"Error saving configuration: {ex.Message}", LogLevel.Error);
                return false;
            }
        }

        /// <summary>
        /// Run the InvoiceMailer with specified options
        /// </summary>
        private async Task<bool> RunInvoiceMailer(bool dryRun, string? invoicesFolder, string? recipientsFile)
        {
            try
            {
                // Prepare arguments for InvoiceMailer
                var args = new List<string>();
                if (dryRun)
                {
                    args.Add("--dry-run");
                    Log("Running in dry-run mode - emails will not be sent", LogLevel.Info);
                }
                else
                {
                    Log("Running in production mode - REAL EMAILS WILL BE SENT", LogLevel.Warning);
                }

                if (!string.IsNullOrEmpty(invoicesFolder))
                {
                    args.Add("--invoices");
                    args.Add(invoicesFolder);
                    Log($"Using custom invoices folder: {invoicesFolder}", LogLevel.Info);
                }

                if (!string.IsNullOrEmpty(recipientsFile))
                {
                    args.Add("--recipients");
                    args.Add(recipientsFile);
                    Log($"Using custom recipients file: {recipientsFile}", LogLevel.Info);
                }

                // Build and invoke the command
                var rootCommand = BuildRootCommand();
                int exitCode = await rootCommand.InvokeAsync(args.ToArray());
                
                if (exitCode == 0)
                {
                    return true;
                }
                else
                {
                    Log($"Process completed with exit code: {exitCode}", LogLevel.Warning);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log($"Error running InvoiceMailer: {ex.Message}", LogLevel.Error);
                return false;
            }
        }

        // Build a command line interface similar to the one in InvoiceMailer
        private RootCommand BuildRootCommand()
        {
            var rootCommand = new RootCommand("InvoiceMailer - Process invoices and send emails");

            // Add command line options
            var dryRunOption = new Option<bool>(
                new[] { "--dry-run" },
                "Run in simulation mode without actually sending emails"
            );

            var invoicesFolderOption = new Option<string>(
                new[] { "--invoices" },
                "Override the invoices folder path"
            );

            var recipientsFileOption = new Option<string>(
                new[] { "--recipients" },
                "Override the recipients CSV file path"
            );

            var checkOption = new Option<bool>(
                new[] { "--check" },
                "Run a health check validation only"
            );

            // Add options to the root command
            rootCommand.AddOption(dryRunOption);
            rootCommand.AddOption(invoicesFolderOption);
            rootCommand.AddOption(recipientsFileOption);
            rootCommand.AddOption(checkOption);

            return rootCommand;
        }
    }
} 