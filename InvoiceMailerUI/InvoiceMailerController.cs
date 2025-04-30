using System;
using System.Collections.Generic;
using System.CommandLine;
using System.IO;
using System.Threading.Tasks;
using Microsoft.Extensions.Configuration;
using Spectre.Console;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Identity.Client;
using Azure.Identity;

namespace InvoiceMailerUI
{
    /// <summary>
    /// Controller for InvoiceMailer operations, separating business logic from UI
    /// </summary>
    public class InvoiceMailerController
    {
        private IConfiguration _configuration;
        private Action<string, LogLevel>? _logCallback;
        private IEmailSender? _emailSender;

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
            LoadConfiguration();
        }

        /// <summary>
        /// Reload configuration from appsettings.json
        /// </summary>
        private void LoadConfiguration()
        {
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
                    
                    if (string.IsNullOrWhiteSpace(configTenantId) || 
                        string.IsNullOrWhiteSpace(configClientId))
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
                "[bold][yellow]⚠️ INTERACTIVE AUTHENTICATION:[/][/]\n\n" +
                "The application now uses interactive authentication with Microsoft Identity Platform.\n" +
                "[green]• When you run the application, it will open a browser window[/]\n" +
                "[green]• You will need to sign in with your organization email[/]\n" +
                "[green]• After signing in, the app will automatically use your email as the sender[/]\n" +
                "[green]• You can override the sender email address during runtime if needed[/]\n\n" +
                "[bold]Required Information:[/]\n" +
                "You'll need to provide your Azure AD Tenant ID and Client ID from your registered application.\n" +
                "No client secret is needed as we're using user-based authentication."
            ).Border(BoxBorder.Rounded)
            .Padding(1, 1);
            
            AnsiConsole.Write(panel);
            AnsiConsole.WriteLine();
            
            // Prompt for settings - PHASE 1: GATHER ALL USER INPUT
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
            
            // Get current or default invoice pattern
            string currentPattern = _configuration.GetValue<string>("InvoiceScanner:DefaultPattern") ?? @"INV\d+";
            
            string wizardInvoicePattern = AnsiConsole.Prompt(
                new TextPrompt<string>($"Enter the invoice key pattern (e.g. INV\\d+, INVOICE_\\d{{4}}, etc):")
                    .DefaultValue(currentPattern)
                    .PromptStyle("green")
                    .ValidationErrorMessage("[red]Invalid regular expression pattern[/]")
                    .Validate(pattern => 
                    {
                        if (string.IsNullOrWhiteSpace(pattern))
                            return ValidationResult.Error("[red]The pattern cannot be empty[/]");
                            
                        // Test if it's a valid regex pattern
                        try 
                        {
                            var _ = new System.Text.RegularExpressions.Regex(pattern);
                            return ValidationResult.Success();
                        }
                        catch (Exception)
                        {
                            return ValidationResult.Error("[red]Invalid regular expression pattern[/]");
                        }
                    }));
            
            AnsiConsole.WriteLine();
            
            // PHASE 2: PROCESSING - Now that we have all user input, do the processing
            bool success = false;
            
            // Use a status spinner for processing
            await AnsiConsole.Status()
                .StartAsync("Saving configuration...", async ctx => 
                {
                    ctx.Spinner(Spinner.Known.Dots);
                    success = await SaveSetupConfiguration(wizardTenantId, wizardClientId, wizardInvoicePattern);
                });
            
            // PHASE 3: RESULTS
            if (success)
            {
                AnsiConsole.MarkupLine("[green]Setup completed successfully![/]");
                AnsiConsole.MarkupLine("[blue]Configuration saved to appsettings.json[/]");
                AnsiConsole.MarkupLine("[yellow]When you first run the application, you'll be asked to sign in with your browser.[/]");
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
        private async Task<bool> SaveSetupConfiguration(string wizardTenantId, string wizardClientId, string invoiceKeyPattern = "INV\\d+")
        {
            Log("Saving configuration from setup wizard...", LogLevel.Info);
            
            try
            {
                string appSettingsPath = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");
                
                // Load existing config if it exists, or create a new one
                JsonNode configRoot;
                string? existingDefaultSenderEmail = null;
                
                if (File.Exists(appSettingsPath))
                {
                    string existingJson = await File.ReadAllTextAsync(appSettingsPath);
                    configRoot = JsonNode.Parse(existingJson) ?? new JsonObject();
                    
                    // Preserve any existing DefaultSenderEmail
                    if (configRoot["Email"] is JsonObject existingEmail)
                    {
                        if (existingEmail.ContainsKey("DefaultSenderEmail"))
                        {
                            existingDefaultSenderEmail = existingEmail["DefaultSenderEmail"]?.GetValue<string>();
                        }
                    }
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
                            ["DefaultPattern"] = invoiceKeyPattern,
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
                emailSection["AuthMode"] = "Interactive";
                
                // Restore the DefaultSenderEmail if it existed
                if (!string.IsNullOrEmpty(existingDefaultSenderEmail))
                {
                    emailSection["DefaultSenderEmail"] = existingDefaultSenderEmail;
                }
                
                // Remove client secret if it exists since we're using interactive login
                if (emailSection.ContainsKey("ClientSecret"))
                {
                    emailSection.Remove("ClientSecret");
                }
                
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
                
                // Reload configuration after saving
                LoadConfiguration();
                Log("Configuration reloaded from disk.", LogLevel.Info);
                
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
        public async Task<bool> RunDryRun(string? invoicesFolder = null, string? recipientsFile = null, string? senderEmail = null)
        {
            Log("Starting dry-run mode...", LogLevel.Info);
            
            try
            {
                bool result = await RunInvoiceMailer(true, invoicesFolder, recipientsFile, senderEmail);
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
        public async Task<bool> RunRealSend(string? invoicesFolder = null, string? recipientsFile = null, string? senderEmail = null)
        {
            Log("Starting real-send mode...", LogLevel.Info);
            Log("WARNING: This will send actual emails to recipients!", LogLevel.Warning);
            
            try
            {
                bool result = await RunInvoiceMailer(false, invoicesFolder, recipientsFile, senderEmail);
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
        /// Verify that all required configuration settings are present
        /// </summary>
        public bool VerifyConfiguration()
        {
            Log("Verifying configuration...", LogLevel.Info);
            
            bool isValid = true;
            
            // Check for required settings
            var emailConfig = _configuration.GetSection("Email");
            var tenantId = emailConfig.GetValue<string>("TenantId") ?? string.Empty;
            var clientId = emailConfig.GetValue<string>("ClientId") ?? string.Empty;

            // Tenant ID is required
            if (string.IsNullOrWhiteSpace(tenantId))
            {
                Log("TenantId is not configured in appsettings.json", LogLevel.Error);
                isValid = false;
            }
            
            // Client ID is required
            if (string.IsNullOrWhiteSpace(clientId))
            {
                Log("ClientId is not configured in appsettings.json", LogLevel.Error);
                isValid = false;
            }
            
            // No need to check for client secret as it's no longer used with interactive auth

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
            
            // Report status
            if (isValid)
            {
                Log("Configuration verification completed successfully", LogLevel.Success);
            }
            else
            {
                Log("Configuration verification failed, please fix the reported issues", LogLevel.Error);
            }
            
            return isValid;
        }

        /// <summary>
        /// Get the absolute path of the invoices folder
        /// </summary>
        public string GetInvoicesFolderPath()
        {
            var scanPath = _configuration.GetValue<string>("InvoiceScanner:ScanPath") ?? "invoices";
            
            // Handle relative or absolute paths
            if (Path.IsPathRooted(scanPath))
            {
                return scanPath;
            }
            else
            {
                return Path.Combine(Directory.GetCurrentDirectory(), scanPath);
            }
        }

        /// <summary>
        /// Get the absolute path of the recipients file
        /// </summary>
        public string GetRecipientsFilePath()
        {
            var csvPath = _configuration.GetValue<string>("RecipientLookup:CsvPath") ?? "recipients.csv";
            
            // Handle relative or absolute paths
            if (Path.IsPathRooted(csvPath))
            {
                return csvPath;
            }
            else
            {
                return Path.Combine(Directory.GetCurrentDirectory(), csvPath);
            }
        }

        /// <summary>
        /// Save configuration to appsettings.json
        /// </summary>
        public bool SaveConfiguration(bool testMode, string tenantId, string clientId, string defaultSenderEmail = "", 
            string? invoicesFolderPath = null, string? recipientsFilePath = null, string? invoiceKeyPattern = null,
            bool useExcelForRecipients = false)
        {
            Log("Saving configuration...", LogLevel.Info);
            
            try
            {
                string appSettingsPath = Path.Combine(Directory.GetCurrentDirectory(), "appsettings.json");
                
                // Load existing file
                string existingJson = File.ReadAllText(appSettingsPath);
                JsonNode? configRoot = JsonNode.Parse(existingJson);
                
                if (configRoot != null)
                {
                    // Update email settings
                    if (configRoot["Email"] is JsonObject emailSection)
                    {
                        emailSection["TenantId"] = tenantId;
                        emailSection["ClientId"] = clientId;
                        emailSection["TestMode"] = testMode;
                        emailSection["AuthMode"] = "Interactive";
                        
                        // Remove client secret if it exists since we're using interactive login
                        if (emailSection.ContainsKey("ClientSecret"))
                        {
                            emailSection.Remove("ClientSecret");
                        }
                        
                        // Remove old SenderEmail key if it exists
                        if (emailSection.ContainsKey("SenderEmail"))
                        {
                            emailSection.Remove("SenderEmail");
                        }
                        
                        // Set or remove DefaultSenderEmail based on the parameter
                        if (!string.IsNullOrWhiteSpace(defaultSenderEmail))
                        {
                            emailSection["DefaultSenderEmail"] = defaultSenderEmail;
                            Log($"Default sender email set to: {defaultSenderEmail}", LogLevel.Info);
                        }
                        else if (emailSection.ContainsKey("DefaultSenderEmail"))
                        {
                            emailSection.Remove("DefaultSenderEmail");
                            Log("Default sender email removed from configuration", LogLevel.Info);
                        }
                    }
                    
                    // Update invoices folder path if provided
                    if (configRoot["InvoiceScanner"] is JsonObject scannerSection)
                    {
                        // Update invoice key pattern if provided
                        if (invoiceKeyPattern != null)
                        {
                            scannerSection["DefaultPattern"] = invoiceKeyPattern;
                            Log($"Updated invoice key pattern to: {invoiceKeyPattern}", LogLevel.Info);
                        }
                        
                        // Update invoices folder path if provided
                        if (invoicesFolderPath != null)
                        {
                            // Convert to relative path if possible
                            string currentDirectory = Directory.GetCurrentDirectory();
                            string path = invoicesFolderPath;
                            
                            if (Path.IsPathRooted(path) && path.StartsWith(currentDirectory))
                            {
                                path = Path.GetRelativePath(currentDirectory, path);
                            }
                            
                            scannerSection["ScanPath"] = path;
                            Log($"Updated invoices folder path to: {path}", LogLevel.Info);
                        }
                    }
                    
                    // Update recipients file path if provided
                    if (recipientsFilePath != null && configRoot["RecipientLookup"] is JsonObject recipientsSection)
                    {
                        // Convert to relative path if possible
                        string currentDirectory = Directory.GetCurrentDirectory();
                        string path = recipientsFilePath;
                        
                        if (Path.IsPathRooted(path) && path.StartsWith(currentDirectory))
                        {
                            path = Path.GetRelativePath(currentDirectory, path);
                        }
                        
                        recipientsSection["CsvPath"] = path;
                        Log($"Updated recipients file path to: {path}", LogLevel.Info);
                        
                        // If using Excel, update the Excel path as well
                        if (useExcelForRecipients && path.EndsWith(".xlsx", StringComparison.OrdinalIgnoreCase))
                        {
                            recipientsSection["ExcelPath"] = path;
                            Log($"Updated Excel recipients file path to: {path}", LogLevel.Info);
                        }
                        else if (useExcelForRecipients)
                        {
                            // Convert CSV path to Excel path if needed
                            string excelPath = Path.ChangeExtension(path, ".xlsx");
                            recipientsSection["ExcelPath"] = excelPath;
                            Log($"Updated Excel recipients file path to: {excelPath}", LogLevel.Info);
                        }
                    }
                    
                    // Save the updated configuration
                    var options = new JsonSerializerOptions { WriteIndented = true };
                    string updatedJson = configRoot.ToJsonString(options);
                    File.WriteAllText(appSettingsPath, updatedJson);
                
                    Log("Configuration saved successfully.", LogLevel.Success);
                    
                    // Reload configuration after saving
                    LoadConfiguration();
                    Log("Configuration reloaded from disk.", LogLevel.Info);
                    
                    return true;
                }
                else
                {
                    Log("Failed to parse appsettings.json.", LogLevel.Error);
                    return false;
                }
            }
            catch (Exception ex)
            {
                Log($"Error saving configuration: {ex.Message}", LogLevel.Error);
                return false;
            }
        }

        /// <summary>
        /// Authenticates the user and initializes the appropriate email sender.
        /// Must be called before RunInvoiceMailer.
        /// </summary>
        /// <param name="dryRun">Indicates if running in dry-run mode.</param>
        /// <returns>True if authentication succeeds, false otherwise.</returns>
        public async Task<bool> AuthenticateUser(bool dryRun)
        {
            Log("Initializing Email Sender...", LogLevel.Info);

            // Get configuration for sender
            var emailConfig = _configuration.GetSection("Email");
            var tenantId = emailConfig.GetValue<string>("TenantId") ?? string.Empty;
            var clientId = emailConfig.GetValue<string>("ClientId") ?? string.Empty;
            var testMode = dryRun || emailConfig.GetValue<bool>("TestMode"); // dryRun forces test mode

            // Check for required configuration needed for authentication
            if (!testMode && (string.IsNullOrWhiteSpace(tenantId) || string.IsNullOrWhiteSpace(clientId)))
            {
                Log("Missing required Email configuration (TenantId, ClientId) for authentication.", LogLevel.Error);
                Log("Please run configuration or use dry-run mode.", LogLevel.Error);
                return false;
            }

            // Create email sender (real or mock)
            if (testMode)
            {
                _emailSender = new MockEmailSender();
                Log("Using MockEmailSender (test mode). Authentication skipped.", LogLevel.Info);
                return true; // No real authentication needed for mock
            }
            else
            {
                Log("Using real EmailSender with Microsoft Graph API.", LogLevel.Info);
                _emailSender = new EmailSender(tenantId, clientId);
                
                // Authenticate user - THIS IS THE INTERACTIVE PART
                Log("Authenticating user... (A browser window may open for you to sign in)", LogLevel.Info);
                try 
                {
                    // Pass silent: false for interactive login
                    bool authenticated = await _emailSender.AuthenticateAsync(silent: false); 
                    if (!authenticated)
                    {
                        Log("Authentication failed. Unable to proceed.", LogLevel.Error);
                        Log("Ensure you have an active internet connection and valid Azure AD credentials.", LogLevel.Error);
                        _emailSender = null; // Clear sender on failure
                        return false;
                    }
                    
                    Log($"Successfully authenticated as {_emailSender.GetUserEmail()}", LogLevel.Success);
                    return true;
                }
                catch (Exception authEx)
                {
                    Log("Authentication failed with an error:", LogLevel.Error);
                    Log(authEx.Message, LogLevel.Error);
                    if (authEx.InnerException != null)
                    {
                        Log($"Inner Exception: {authEx.InnerException.Message}", LogLevel.Error);
                    }
                    Log("Ensure you have an active internet connection and valid Azure AD credentials.", LogLevel.Error);
                     _emailSender = null; // Clear sender on failure
                    return false;
                }
            }
        }

        /// <summary>
        /// Run the InvoiceMailer with specified options. Assumes AuthenticateUser has been called successfully.
        /// </summary>
        private async Task<bool> RunInvoiceMailer(bool dryRun, string? invoicesFolder, string? recipientsFile, string? senderEmail = null)
        {
            Log("Starting Invoice Mailer process...", LogLevel.Info);
            
            // Ensure email sender is initialized (should be done by AuthenticateUser)
            if (_emailSender == null)
            {
                Log("Email sender is not initialized. Authentication might have failed or was not performed.", LogLevel.Error);
                return false;
            }
            
            try
            {
                // Configuration is already checked during Authentication setup if not in test mode
                var emailConfig = _configuration.GetSection("Email");
                
                // *** Authentication is now done BEFORE this method is called ***
                // *** Remove the authentication block from here ***
                
                // Check if there's a default sender email in the configuration
                var defaultSenderEmail = emailConfig.GetValue<string>("DefaultSenderEmail") ?? string.Empty;
                
                // Set default sender email if one is configured
                if (!string.IsNullOrEmpty(defaultSenderEmail))
                {
                    _emailSender.SetSenderEmail(defaultSenderEmail);
                    Log($"Using default sender email from configuration: {defaultSenderEmail}", LogLevel.Info);
                }
                
                // Apply sender email override if provided by the user at runtime
                if (!string.IsNullOrEmpty(senderEmail))
                {
                    _emailSender.SetSenderEmail(senderEmail);
                    Log($"Sender email overridden by user to {senderEmail}", LogLevel.Info);
                }
                else
                {
                    // If no user override, make sure the sender is at least the authenticated user or default
                    // (This ensures sender is set even if default wasn't configured)
                    var currentEffectiveSender = _emailSender.GetEffectiveSenderEmail();
                    if(string.IsNullOrEmpty(currentEffectiveSender) && !(_emailSender is MockEmailSender))
                    {
                         _emailSender.SetSenderEmail(_emailSender.GetUserEmail());
                         Log($"Using authenticated user email as sender: {_emailSender.GetUserEmail()}", LogLevel.Info);
                    }
                }

                var effectiveSender = _emailSender.GetEffectiveSenderEmail();
                if (string.IsNullOrEmpty(effectiveSender))
                {
                     Log("Effective sender email could not be determined. Cannot proceed.", LogLevel.Error);
                     return false;
                }
                Log($"Effective sender for this run: {effectiveSender}", LogLevel.Info);

                
                // Prepare the invoice scanner
                // First, try to use the parameter if provided
                var scanPath = invoicesFolder;
                
                // If parameter is null or empty, fetch from config
                if (string.IsNullOrWhiteSpace(scanPath))
                {
                    scanPath = _configuration.GetValue<string>("InvoiceScanner:ScanPath");
                    
                    // If still null, use default
                    if (string.IsNullOrWhiteSpace(scanPath))
                    {
                        scanPath = "invoices";
                        Log("No invoices folder configured. Using default: 'invoices'", LogLevel.Warning);
                    }
                    else
                    {
                        Log($"Using configured invoices folder from settings: '{scanPath}'", LogLevel.Info);
                    }
                }
                else
                {
                    Log($"Using user-provided invoices folder: '{scanPath}'", LogLevel.Info);
                }
                
                var pattern = _configuration.GetValue<string>("InvoiceScanner:DefaultPattern") ?? @"INV\d+";
                var caseInsensitive = _configuration.GetValue<bool>("InvoiceScanner:CaseInsensitive");
                
                // Debug log the actual paths being used
                Log($"DEBUG: Invoice folder path being used: '{scanPath}'", LogLevel.Info);
                
                // Prepare the recipient lookup
                // First, try to use the parameter if provided
                var csvPath = recipientsFile;
                
                // If parameter is null or empty, fetch from config
                if (string.IsNullOrWhiteSpace(csvPath))
                {
                    csvPath = _configuration.GetValue<string>("RecipientLookup:CsvPath");
                    
                    // If still null, use default
                    if (string.IsNullOrWhiteSpace(csvPath))
                    {
                        csvPath = "recipients.csv";
                        Log("No recipients file configured. Using default: 'recipients.csv'", LogLevel.Warning);
                    }
                    else
                    {
                        Log($"Using configured recipients file from settings: '{csvPath}'", LogLevel.Info);
                    }
                }
                else
                {
                    Log($"Using user-provided recipients file: '{csvPath}'", LogLevel.Info);
                }
                
                // Debug log the actual paths being used
                Log($"DEBUG: Recipients file path being used: '{csvPath}'", LogLevel.Info);
                
                // Check if directories and files exist
                if (!Directory.Exists(scanPath))
                {
                    Log($"WARNING: Invoices folder '{scanPath}' does not exist", LogLevel.Warning);
                }
                
                if (!File.Exists(csvPath))
                {
                    Log($"WARNING: Recipients file '{csvPath}' does not exist", LogLevel.Warning);
                }
                
                Log($"Scanning for invoices in '{scanPath}' with pattern '{pattern}'", LogLevel.Info);
                
                // In dry-run mode, display the pattern being used
                if (dryRun)
                {
                    Log($"Using invoice key pattern: {pattern}", LogLevel.Info);
                }
                
                var invoiceScanner = new InvoiceScanner(scanPath, pattern, caseInsensitive);
                
                Log($"Loading recipients from '{csvPath}'", LogLevel.Info);
                var recipientLookup = new RecipientLookup(csvPath);
                
                // Find all invoices
                var invoiceFiles = invoiceScanner.ScanForInvoices().ToList();
                Log($"Found {invoiceFiles.Count} invoice files", LogLevel.Info);
                
                if (invoiceFiles.Count == 0)
                {
                    Log("No invoice files found matching the pattern. Nothing to process.", LogLevel.Warning);
                    return true; // Technically successful, just nothing to do
                }
                
                // Display the sender email information before processing
                Log($"Emails will be sent from: {effectiveSender}", LogLevel.Info);
                
                // Process each invoice
                int successCount = 0;
                int failureCount = 0;
                
                foreach (var (invoicePath, invoiceId) in invoiceFiles)
                {
                    var invoiceFile = Path.GetFileName(invoicePath);
                    
                    if (string.IsNullOrEmpty(invoiceId))
                    {
                        Log($"Could not extract invoice ID from '{invoiceFile}', skipping", LogLevel.Warning);
                        failureCount++;
                        continue;
                    }
                    
                    // Look up recipient email
                    var recipientEmail = recipientLookup.GetEmail(invoiceId);
                    if (string.IsNullOrEmpty(recipientEmail))
                    {
                        Log($"No recipient email found for invoice ID '{invoiceId}' in '{csvPath}', skipping", LogLevel.Warning);
                        failureCount++;
                        continue;
                    }
                    
                    // Send the email
                    try
                    {
                        Log($"Processing invoice '{invoiceId}' for '{recipientEmail}'...", LogLevel.Info);
                        
                        await _emailSender.SendEmailAsync(
                            recipientEmail,
                            $"Invoice {invoiceId}",
                            $"Please find attached invoice {invoiceId}. " +
                            "If you have any questions, please contact our accounting department.",
                            invoicePath);
                        
                        Log($"Successfully {(dryRun ? "simulated sending" : "sent")} invoice '{invoiceId}' to '{recipientEmail}'", LogLevel.Success);
                        successCount++;
                    }
                    catch (Exception ex)
                    {
                        Log($"Failed to send invoice '{invoiceId}' to '{recipientEmail}': {ex.Message}", LogLevel.Error);
                        if (ex.InnerException != null)
                        {
                            Log($"Reason: {ex.InnerException.Message}", LogLevel.Error);
                        }
                        failureCount++;
                    }
                }
                
                // Report results
                Log($"Completed processing {invoiceFiles.Count} invoices.", LogLevel.Info);
                Log($"Success: {successCount}, Failures: {failureCount}", 
                    failureCount > 0 ? LogLevel.Warning : LogLevel.Success);
                
                return failureCount == 0;
            }
            catch (Exception ex)
            {
                Log($"Error during Invoice Mailer run: {ex.Message}", LogLevel.Error);
                if (ex.InnerException != null)
                {
                    Log($"Details: {ex.InnerException.Message}", LogLevel.Error);
                }
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

        /// <summary>
        /// Returns the Email configuration section as a dictionary
        /// </summary>
        public Dictionary<string, string> GetEmailConfiguration()
        {
            var result = new Dictionary<string, string>();
            var emailConfig = _configuration.GetSection("Email");
            
            foreach (var child in emailConfig.GetChildren())
            {
                result[child.Key] = child.Value ?? string.Empty;
            }
            
            return result;
        }

        /// <summary>
        /// Gets the current invoice key pattern from configuration
        /// </summary>
        public string GetInvoiceKeyPattern()
        {
            var pattern = _configuration.GetValue<string>("InvoiceScanner:DefaultPattern");
            return string.IsNullOrWhiteSpace(pattern) ? @"INV\d+" : pattern;
        }
    }
} 