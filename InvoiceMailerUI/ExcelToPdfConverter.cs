using QuestPDF.Fluent;
using QuestPDF.Helpers;
using QuestPDF.Infrastructure;
using ClosedXML.Excel;
using Serilog;

namespace InvoiceMailerUI
{
    public class ExcelToPdfConverter
    {
        private readonly ILogger _logger;
        
        public ExcelToPdfConverter()
        {
            _logger = Log.ForContext<ExcelToPdfConverter>();
            // Register QuestPDF license (Community Edition)
            QuestPDF.Settings.License = LicenseType.Community;
        }
        
        /// <summary>
        /// Converts an Excel file to PDF format
        /// </summary>
        /// <param name="excelFilePath">Path to the Excel file to convert</param>
        /// <returns>Path to the generated PDF file (in temp directory)</returns>
        public string ConvertToPdf(string excelFilePath)
        {
            _logger.Information("Converting Excel file to PDF: {FilePath}", excelFilePath);
            
            try
            {
                using var workbook = new XLWorkbook(excelFilePath);
                var worksheet = workbook.Worksheets.FirstOrDefault();
                
                if (worksheet == null)
                {
                    throw new InvalidOperationException("No worksheets found in the Excel file");
                }
                
                string fileName = Path.GetFileNameWithoutExtension(excelFilePath);
                string outputFilePath = Path.Combine(Path.GetTempPath(), $"{fileName}.pdf");
                
                // Create PDF document with Excel-like styling
                Document.Create(container =>
                {
                    container.Page(page =>
                    {
                        page.Size(PageSizes.A4);
                        page.Margin(20);
                        
                        page.Header().Text($"Invoice: {fileName}")
                            .FontSize(14)
                            .SemiBold()
                            .FontColor(Colors.Blue.Medium);
                            
                        page.Content().Element(ComposeContent);
                        
                        page.Footer().AlignCenter().Text(text =>
                        {
                            text.Span("Page ");
                            text.CurrentPageNumber();
                            text.Span(" of ");
                            text.TotalPages();
                        });
                    });
                    
                    // Compose the worksheet content
                    void ComposeContent(IContainer container)
                    {
                        container.Table(table =>
                        {
                            // Get used range
                            var usedRange = worksheet.RangeUsed();
                            if (usedRange == null)
                            {
                                table.Cell().Text("No data found in worksheet");
                                return;
                            }
                            
                            int startRow = usedRange.FirstRow().RowNumber();
                            int endRow = usedRange.LastRow().RowNumber();
                            int startColumn = usedRange.FirstColumn().ColumnNumber();
                            int endColumn = usedRange.LastColumn().ColumnNumber();
                            
                            // Create columns
                            for (int col = startColumn; col <= endColumn; col++)
                            {
                                table.ColumnsDefinition(columns =>
                                {
                                    columns.RelativeColumn();
                                });
                            }
                            
                            // Add header row
                            table.Header(header =>
                            {
                                for (int col = startColumn; col <= endColumn; col++)
                                {
                                    var cellValue = worksheet.Cell(startRow, col).GetString();
                                    
                                    header.Cell().Border(1).Background(Colors.Grey.Lighten2).Text(cellValue)
                                        .SemiBold()
                                        .FontColor(Colors.Black);
                                }
                            });
                            
                            // Add data rows
                            for (int row = startRow + 1; row <= endRow; row++)
                            {
                                for (int col = startColumn; col <= endColumn; col++)
                                {
                                    var cellValue = worksheet.Cell(row, col).GetString();
                                    
                                    table.Cell().Border(1).BorderColor(Colors.Grey.Lighten3).Text(cellValue);
                                }
                            }
                            
                            table.ExtendLastCellsToTableBottom();
                        });
                    }
                })
                .GeneratePdf(outputFilePath);
                
                _logger.Information("Successfully converted Excel to PDF: {OutputPath}", outputFilePath);
                return outputFilePath;
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error converting Excel to PDF: {FilePath}", excelFilePath);
                throw;
            }
        }
    }
} 