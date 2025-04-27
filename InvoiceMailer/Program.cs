using Microsoft.Extensions.Configuration;
using Serilog;
using Serilog.Events;
using InvoiceMailer;
using System.IO;

// Load configuration
var configuration = new ConfigurationBuilder()
    .SetBasePath(Directory.GetCurrentDirectory())
    .AddJsonFile("appsettings.json", optional: false, reloadOnChange: true)
    .Build();

// Configure Serilog
Log.Logger = new LoggerConfiguration()
    .ReadFrom.Configuration(configuration)
    .CreateLogger();

try
{
    Log.Information("Starting InvoiceMailer application");
    
    // Get application settings from config
    var appSettings = configuration.GetSection("ApplicationSettings");
    var appName = appSettings.GetValue<string>("ApplicationName");
    var environment = appSettings.GetValue<string>("Environment");
    
    Log.Information("Application {AppName} running in {Environment} environment", appName, environment);
    
    // Get InvoiceScanner configuration
    var scannerConfig = configuration.GetSection("InvoiceScanner");
    var scanPath = scannerConfig.GetValue<string>("ScanPath") ?? "invoices";
    var patternOverride = scannerConfig.GetValue<string>("DefaultPattern") ?? @"INV\d+"; // Provide default to avoid null
    var caseInsensitive = scannerConfig.GetValue<bool>("CaseInsensitive");
    
    // Create full path for scan folder
    var scanFolder = Path.IsPathRooted(scanPath)
        ? scanPath
        : Path.Combine(Directory.GetCurrentDirectory(), scanPath);
    
    // Create the invoices folder if it doesn't exist
    if (!Directory.Exists(scanFolder))
    {
        Directory.CreateDirectory(scanFolder);
        Log.Information("Created directory: {Directory}", scanFolder);
    }
    
    // Get RecipientLookup configuration
    var recipientConfig = configuration.GetSection("RecipientLookup");
    var csvPath = recipientConfig.GetValue<string>("CsvPath") ?? "recipients.csv";
    
    // Create full path for recipient CSV
    var recipientsPath = Path.IsPathRooted(csvPath)
        ? csvPath
        : Path.Combine(Directory.GetCurrentDirectory(), csvPath);
    
    // Initialize recipient lookup
    var recipientLookup = new RecipientLookup(recipientsPath);
    
    // Scan for invoices using configuration
    var scanner = new InvoiceScanner(scanFolder, patternOverride, caseInsensitive);
    var invoices = scanner.ScanForInvoices().ToList();
    
    Log.Information("Found {Count} invoice files with valid keys", invoices.Count);
    
    // Display results with recipient information
    if (invoices.Any())
    {
        Console.WriteLine("\n--- Detected Invoice Files ---");
        foreach (var (path, key) in invoices)
        {
            var email = recipientLookup.GetEmail(key);
            string recipient = email != null ? email : "NO RECIPIENT FOUND";
            
            Console.WriteLine($"Invoice: {key} - File: {Path.GetFileName(path)} - Recipient: {recipient}");
        }
        Console.WriteLine();
    }
    else
    {
        Console.WriteLine("\nNo invoice files with valid keys were found.\n");
    }
    
    Log.Information("InvoiceMailer application completed successfully");
    return 0;
}
catch (Exception ex)
{
    Log.Fatal(ex, "InvoiceMailer application terminated unexpectedly");
    return 1;
}
finally
{
    // Ensure to flush and close the log
    Log.CloseAndFlush();
}
