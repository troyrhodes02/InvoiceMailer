using Serilog;

namespace InvoiceMailerUI
{
    public class RecipientLookup
    {
        private readonly ILogger _logger;
        private readonly Dictionary<string, string> _recipients = new();
        private readonly string _csvPath;

        public RecipientLookup(string csvPath)
        {
            _logger = Log.ForContext<RecipientLookup>();
            _csvPath = csvPath;
            
            LoadRecipientsFile();
        }

        private void LoadRecipientsFile()
        {
            if (!File.Exists(_csvPath))
            {
                _logger.Error("Recipients file not found: {FilePath}", _csvPath);
                return;
            }

            try
            {
                _logger.Information("Loading recipients from: {FilePath}", _csvPath);
                
                // Skip header row if it exists
                bool isFirstLine = true;
                
                foreach (var line in File.ReadLines(_csvPath))
                {
                    // Skip empty lines
                    if (string.IsNullOrWhiteSpace(line))
                        continue;
                    
                    // Check if first line is a header and skip if needed
                    if (isFirstLine)
                    {
                        isFirstLine = false;
                        if (line.Contains("Key") && line.Contains("Email"))
                            continue;
                    }
                    
                    // Split the line into key and email
                    var parts = line.Split(',');
                    if (parts.Length >= 2)
                    {
                        string key = parts[0].Trim();
                        string email = parts[1].Trim();
                        
                        if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(email))
                        {
                            _recipients[key] = email;
                            _logger.Debug("Added recipient mapping: {Key} -> {Email}", key, email);
                        }
                        else
                        {
                            _logger.Warning("Invalid recipient data in line: {Line}", line);
                        }
                    }
                    else
                    {
                        _logger.Warning("Malformed CSV line: {Line}", line);
                    }
                }
                
                _logger.Information("Loaded {Count} recipient mappings", _recipients.Count);
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error loading recipients file: {FilePath}", _csvPath);
            }
        }

        public string? GetEmail(string key)
        {
            if (_recipients.TryGetValue(key, out var email))
            {
                _logger.Information("Found recipient for invoice {Key}: {Email}", key, email);
                return email;
            }
            
            _logger.Warning("No recipient found for invoice {Key}", key);
            return null;
        }
        
        public int Count => _recipients.Count;
    }
} 