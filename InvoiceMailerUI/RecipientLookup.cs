using Serilog;
using ClosedXML.Excel;

namespace InvoiceMailerUI
{
    public class RecipientLookup
    {
        private readonly ILogger _logger;
        private readonly Dictionary<string, string> _recipients = new();
        private readonly string _filePath;
        private readonly bool _isExcelFile;

        public RecipientLookup(string filePath)
        {
            _logger = Log.ForContext<RecipientLookup>();
            _filePath = filePath;
            _isExcelFile = Path.GetExtension(filePath).ToLowerInvariant() == ".xlsx";
            
            LoadRecipientsFile();
        }

        private void LoadRecipientsFile()
        {
            if (!File.Exists(_filePath))
            {
                _logger.Error("Recipients file not found: {FilePath}", _filePath);
                return;
            }

            try
            {
                _logger.Information("Loading recipients from: {FilePath}", _filePath);
                
                if (_isExcelFile)
                {
                    LoadFromExcel();
                }
                else
                {
                    LoadFromCsv();
                }
                
                _logger.Information("Loaded {Count} recipient mappings", _recipients.Count);
                
                // Validate for missing or duplicate keys
                ValidateRecipients();
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error loading recipients file: {FilePath}", _filePath);
            }
        }
        
        private void LoadFromExcel()
        {
            using var workbook = new XLWorkbook(_filePath);
            var worksheet = workbook.Worksheets.FirstOrDefault();
            
            if (worksheet == null)
            {
                _logger.Error("No worksheets found in the Excel file: {FilePath}", _filePath);
                return;
            }
            
            // Find the header row and columns
            int invoiceKeyColumnIndex = -1;
            int emailColumnIndex = -1;
            
            var headerRow = worksheet.FirstRowUsed();
            if (headerRow == null)
            {
                _logger.Error("Excel file is empty: {FilePath}", _filePath);
                return;
            }
            
            // Find column indices based on headers (case-insensitive)
            int columnCount = headerRow.CellsUsed().Count();
            for (int i = 1; i <= columnCount; i++)
            {
                var cellValue = headerRow.Cell(i).GetString().Trim();
                
                if (cellValue.Equals("InvoiceKey", StringComparison.OrdinalIgnoreCase))
                {
                    invoiceKeyColumnIndex = i;
                }
                else if (cellValue.Equals("Email", StringComparison.OrdinalIgnoreCase))
                {
                    emailColumnIndex = i;
                }
            }
            
            if (invoiceKeyColumnIndex == -1 || emailColumnIndex == -1)
            {
                _logger.Error("Required headers not found in Excel file. Need 'InvoiceKey' and 'Email' columns (case insensitive): {FilePath}", _filePath);
                return;
            }
            
            // Process data rows
            foreach (var row in worksheet.RowsUsed().Skip(1)) // Skip header row
            {
                string key = row.Cell(invoiceKeyColumnIndex).GetString().Trim();
                string email = row.Cell(emailColumnIndex).GetString().Trim();
                
                if (!string.IsNullOrEmpty(key) && !string.IsNullOrEmpty(email))
                {
                    _recipients[key] = email;
                    _logger.Debug("Added recipient mapping from Excel: {Key} -> {Email}", key, email);
                }
                else
                {
                    _logger.Warning("Invalid recipient data in Excel row {RowNumber}", row.RowNumber());
                }
            }
        }
        
        private void LoadFromCsv()
        {
            // Skip header row if it exists
            bool isFirstLine = true;
            
            foreach (var line in File.ReadLines(_filePath))
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
                        _logger.Debug("Added recipient mapping from CSV: {Key} -> {Email}", key, email);
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
        }
        
        private void ValidateRecipients()
        {
            if (_recipients.Count == 0)
            {
                _logger.Warning("No recipients loaded from file: {FilePath}", _filePath);
            }
            
            var duplicateKeys = _recipients.Keys
                .GroupBy(k => k)
                .Where(g => g.Count() > 1)
                .Select(g => g.Key)
                .ToList();
                
            if (duplicateKeys.Any())
            {
                foreach (var key in duplicateKeys)
                {
                    _logger.Warning("Duplicate invoice key found in recipients file: {Key}", key);
                }
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