using System.Text.RegularExpressions;
using Serilog;

namespace InvoiceMailer;

public class InvoiceScanner
{
    private readonly ILogger _logger;
    private readonly string _folderPath;
    private readonly Regex _keyPattern;

    public InvoiceScanner(string folderPath, string pattern = @"INV\d+", bool caseInsensitive = true)
    {
        _logger = Log.ForContext<InvoiceScanner>();
        _folderPath = folderPath;
        
        var regexOptions = RegexOptions.Compiled;
        if (caseInsensitive)
        {
            regexOptions |= RegexOptions.IgnoreCase;
        }
        
        _keyPattern = new Regex(pattern, regexOptions);
        
        _logger.Debug("InvoiceScanner initialized for folder: {FolderPath} with pattern: {Pattern} (case {CaseSensitivity})", 
            folderPath, pattern, caseInsensitive ? "insensitive" : "sensitive");
    }

    public IEnumerable<(string FilePath, string Key)> ScanForInvoices()
    {
        _logger.Information("Starting invoice scan in folder: {FolderPath}", _folderPath);
        
        if (!Directory.Exists(_folderPath))
        {
            _logger.Error("Folder not found: {FolderPath}", _folderPath);
            yield break;
        }

        var fileExtensions = new[] { "*.pdf", "*.xlsx" };
        
        foreach (var extension in fileExtensions)
        {
            _logger.Debug("Scanning for files with extension: {Extension}", extension);
            
            string[] files = Array.Empty<string>();
            
            try
            {
                files = Directory.GetFiles(_folderPath, extension, SearchOption.TopDirectoryOnly);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error scanning for {Extension} files", extension);
                continue;
            }
            
            foreach (var file in files)
            {
                var fileName = Path.GetFileName(file);
                var match = _keyPattern.Match(fileName);
                
                if (match.Success)
                {
                    var key = match.Value;
                    _logger.Information("Found invoice file: {FileName} with key: {Key}", fileName, key);
                    yield return (file, key);
                }
                else
                {
                    _logger.Debug("No invoice key found in file: {FileName}", fileName);
                }
            }
        }
        
        _logger.Information("Invoice scan completed");
    }
} 