using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using Spectre.Console;

namespace InvoiceMailerUI
{
    /// <summary>
    /// Provides native file and folder browsing capabilities for both Windows and macOS
    /// </summary>
    public static class NativeFileBrowser
    {
        private static readonly bool IsWindows = OperatingSystem.IsWindows();
        private static readonly bool IsMacOS = OperatingSystem.IsMacOS();
        
        /// <summary>
        /// Opens a native folder browser dialog
        /// </summary>
        /// <param name="title">Dialog title</param>
        /// <param name="initialDirectory">Initial directory to display</param>
        /// <returns>Selected folder path or null if canceled</returns>
        public static string? BrowseForFolder(string title, string? initialDirectory = null)
        {
            try
            {
                // Default to current directory if initialDirectory is null or doesn't exist
                if (string.IsNullOrEmpty(initialDirectory) || !Directory.Exists(initialDirectory))
                {
                    initialDirectory = Directory.GetCurrentDirectory();
                }
                
                // Windows: Use System.Windows.Forms.FolderBrowserDialog
                if (IsWindows)
                {
                    return BrowseFolderWindows(title, initialDirectory);
                }
                // macOS: Use NSOpenPanel via AppleScript
                else if (IsMacOS)
                {
                    return BrowseFolderMacOS(title, initialDirectory);
                }
                // Other platforms: Use console-based browser
                else
                {
                    AnsiConsole.MarkupLine("[yellow]Native file browsing not supported on this platform. Using console browser.[/]");
                    return FileBrowser.BrowseForFolder(initialDirectory);
                }
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error opening native folder browser: {ex.Message}[/]");
                AnsiConsole.MarkupLine("[yellow]Falling back to console browser.[/]");
                return FileBrowser.BrowseForFolder(initialDirectory);
            }
        }

        /// <summary>
        /// Opens a native file browser dialog
        /// </summary>
        /// <param name="title">Dialog title</param>
        /// <param name="filter">File filter (e.g., "CSV Files (*.csv)|*.csv")</param>
        /// <param name="initialDirectory">Initial directory to display</param>
        /// <returns>Selected file path or null if canceled</returns>
        public static string? BrowseForFile(string title, string filter, string? initialDirectory = null)
        {
            try
            {
                // Default to current directory if initialDirectory is null or doesn't exist
                if (string.IsNullOrEmpty(initialDirectory) || !Directory.Exists(initialDirectory))
                {
                    initialDirectory = Directory.GetCurrentDirectory();
                }
                
                // Windows: Use System.Windows.Forms.OpenFileDialog
                if (IsWindows)
                {
                    return BrowseFileWindows(title, filter, initialDirectory);
                }
                // macOS: Use NSOpenPanel via AppleScript
                else if (IsMacOS)
                {
                    return BrowseFileMacOS(title, filter, initialDirectory);
                }
                // Other platforms: Use console-based browser
                else
                {
                    AnsiConsole.MarkupLine("[yellow]Native file browsing not supported on this platform. Using console browser.[/]");
                    // Convert filter to console browser format (e.g., "*.csv")
                    string consoleFilter = "*.*";
                    if (filter.Contains("*."))
                    {
                        // Extract file extension pattern
                        int startIndex = filter.IndexOf("*.");
                        if (startIndex >= 0)
                        {
                            int endIndex = filter.IndexOf("|", startIndex);
                            if (endIndex < 0) endIndex = filter.Length;
                            consoleFilter = filter.Substring(startIndex, endIndex - startIndex);
                        }
                    }
                    return FileBrowser.BrowseForFile(initialDirectory, consoleFilter);
                }
            }
            catch (Exception ex)
            {
                AnsiConsole.MarkupLine($"[red]Error opening native file browser: {ex.Message}[/]");
                AnsiConsole.MarkupLine("[yellow]Falling back to console browser.[/]");
                return FileBrowser.BrowseForFile(initialDirectory);
            }
        }

        #region Windows Implementation
        
        private static string? BrowseFolderWindows(string title, string initialDirectory)
        {
            // Check if running on Windows
            if (!IsWindows)
            {
                throw new PlatformNotSupportedException("This method only supports Windows platforms.");
            }
            
            // Dynamically load the Windows Forms assembly
            try
            {
                // Import the required reference
                var assembly = System.Reflection.Assembly.Load("System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089");
                
                // Create FolderBrowserDialog using reflection
                var dialogType = assembly.GetType("System.Windows.Forms.FolderBrowserDialog");
                var dialog = Activator.CreateInstance(dialogType);
                
                // Set properties
                dialogType.GetProperty("Description").SetValue(dialog, title);
                dialogType.GetProperty("SelectedPath").SetValue(dialog, initialDirectory);
                dialogType.GetProperty("ShowNewFolderButton").SetValue(dialog, true);
                
                // Show dialog
                var dialogResult = dialogType.GetMethod("ShowDialog").Invoke(dialog, null);
                
                // Check result
                var resultType = assembly.GetType("System.Windows.Forms.DialogResult");
                var okValue = resultType.GetField("OK").GetValue(null);
                
                if (dialogResult.Equals(okValue))
                {
                    // Get selected path
                    return (string)dialogType.GetProperty("SelectedPath").GetValue(dialog);
                }
                
                return null;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to use Windows Forms dialog: {ex.Message}", ex);
            }
        }

        private static string? BrowseFileWindows(string title, string filter, string initialDirectory)
        {
            // Check if running on Windows
            if (!IsWindows)
            {
                throw new PlatformNotSupportedException("This method only supports Windows platforms.");
            }
            
            // Dynamically load the Windows Forms assembly
            try
            {
                // Import the required reference
                var assembly = System.Reflection.Assembly.Load("System.Windows.Forms, Version=4.0.0.0, Culture=neutral, PublicKeyToken=b77a5c561934e089");
                
                // Create OpenFileDialog using reflection
                var dialogType = assembly.GetType("System.Windows.Forms.OpenFileDialog");
                var dialog = Activator.CreateInstance(dialogType);
                
                // Set properties
                dialogType.GetProperty("Title").SetValue(dialog, title);
                dialogType.GetProperty("InitialDirectory").SetValue(dialog, initialDirectory);
                dialogType.GetProperty("Filter").SetValue(dialog, filter);
                dialogType.GetProperty("CheckFileExists").SetValue(dialog, true);
                
                // Show dialog
                var dialogResult = dialogType.GetMethod("ShowDialog").Invoke(dialog, null);
                
                // Check result
                var resultType = assembly.GetType("System.Windows.Forms.DialogResult");
                var okValue = resultType.GetField("OK").GetValue(null);
                
                if (dialogResult.Equals(okValue))
                {
                    // Get selected file
                    return (string)dialogType.GetProperty("FileName").GetValue(dialog);
                }
                
                return null;
            }
            catch (Exception ex)
            {
                throw new Exception($"Failed to use Windows Forms dialog: {ex.Message}", ex);
            }
        }
        
        #endregion

        #region macOS Implementation
        
        private static string? BrowseFolderMacOS(string title, string initialDirectory)
        {
            // Check if running on macOS
            if (!IsMacOS)
            {
                throw new PlatformNotSupportedException("This method only supports macOS platforms.");
            }
            
            // Normalize directory path for AppleScript
            initialDirectory = initialDirectory.Replace("\\", "/");
            
            // Create a temporary file to store the selected path
            string tempFile = Path.Combine(Path.GetTempPath(), $"folder_select_{Guid.NewGuid()}.txt");
            
            // Replace double quotes with escaped quotes for AppleScript
            string escapedTitle = title.Replace("\"", "\\\"");
            string escapedInitialDirectory = initialDirectory.Replace("\"", "\\\"");
            string escapedTempFile = tempFile.Replace("\\", "/").Replace("\"", "\\\"");
            
            // Build the AppleScript command to show a folder selection dialog
            string appleScript = $"tell application \"Finder\"\n" +
                                $"  set defaultFolder to POSIX file \"{escapedInitialDirectory}\"\n" +
                                $"  try\n" +
                                $"    set selectedFolder to choose folder with prompt \"{escapedTitle}\" default location defaultFolder\n" +
                                $"    set posixPath to POSIX path of selectedFolder\n" +
                                $"    do shell script \"echo \" & quoted form of posixPath & \" > {escapedTempFile}\"\n" +
                                $"  on error\n" +
                                $"    -- Dialog was cancelled\n" +
                                $"  end try\n" +
                                $"end tell";
            
            try
            {
                // Debug the AppleScript for troubleshooting
                AnsiConsole.MarkupLine($"[grey]Executing AppleScript for folder selection[/]");
                
                // Create a temporary script file
                string scriptFile = Path.Combine(Path.GetTempPath(), $"applescript_{Guid.NewGuid()}.scpt");
                File.WriteAllText(scriptFile, appleScript);
                
                // Execute the AppleScript from the file
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "osascript",
                        Arguments = scriptFile,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true
                    }
                };
                
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();
                
                // Clean up the script file
                if (File.Exists(scriptFile))
                {
                    try { File.Delete(scriptFile); } catch { /* Ignore cleanup errors */ }
                }
                
                // If there's an error with AppleScript, log it
                if (!string.IsNullOrEmpty(error))
                {
                    AnsiConsole.MarkupLine($"[red]AppleScript error: {error}[/]");
                }
                
                // Read the selected path from the temporary file (if the dialog wasn't canceled)
                if (File.Exists(tempFile))
                {
                    string selectedPath = File.ReadAllText(tempFile).Trim();
                    File.Delete(tempFile); // Clean up temp file
                    return string.IsNullOrEmpty(selectedPath) ? null : selectedPath;
                }
                
                return null;
            }
            catch (Exception ex)
            {
                if (File.Exists(tempFile))
                {
                    try { File.Delete(tempFile); } catch { /* Ignore cleanup errors */ }
                }
                throw new Exception($"Failed to use macOS file dialog: {ex.Message}", ex);
            }
        }
        
        private static string? BrowseFileMacOS(string title, string filter, string initialDirectory)
        {
            // Check if running on macOS
            if (!IsMacOS)
            {
                throw new PlatformNotSupportedException("This method only supports macOS platforms.");
            }
            
            // Normalize directory path for AppleScript
            initialDirectory = initialDirectory.Replace("\\", "/");
            
            // Extract file extensions from the filter
            string fileTypes = string.Empty;
            if (filter.Contains("*."))
            {
                // Parse the filter string (e.g., "CSV Files (*.csv)|*.csv")
                var parts = filter.Split('|');
                for (int i = 0; i < parts.Length; i++)
                {
                    string part = parts[i];
                    if (part.StartsWith("*."))
                    {
                        fileTypes += $"\"{part.Substring(2)}\","; // Extract "csv" from "*.csv"
                    }
                }
                
                // Remove trailing comma
                if (fileTypes.EndsWith(","))
                {
                    fileTypes = fileTypes.Substring(0, fileTypes.Length - 1);
                }
            }
            
            // Create a temporary file to store the selected path
            string tempFile = Path.Combine(Path.GetTempPath(), $"file_select_{Guid.NewGuid()}.txt");
            
            // Replace double quotes with escaped quotes for AppleScript
            string escapedTitle = title.Replace("\"", "\\\"");
            string escapedInitialDirectory = initialDirectory.Replace("\"", "\\\"");
            string escapedTempFile = tempFile.Replace("\\", "/").Replace("\"", "\\\"");
            
            // Build the AppleScript command to show a file selection dialog
            string fileTypesScript = string.IsNullOrEmpty(fileTypes)
                ? "" // No file type restriction
                : $"  set fileTypes to {{{fileTypes}}}\n";
            
            string fileTypeParam = string.IsNullOrEmpty(fileTypes) 
                ? "" 
                : "of type fileTypes ";
                
            string appleScript = $"tell application \"Finder\"\n" +
                                $"  set defaultFolder to POSIX file \"{escapedInitialDirectory}\"\n" +
                                $"{fileTypesScript}" +
                                $"  try\n" +
                                $"    set selectedFile to choose file with prompt \"{escapedTitle}\" {fileTypeParam}default location defaultFolder\n" +
                                $"    set posixPath to POSIX path of selectedFile\n" +
                                $"    do shell script \"echo \" & quoted form of posixPath & \" > {escapedTempFile}\"\n" +
                                $"  on error\n" +
                                $"    -- Dialog was cancelled\n" +
                                $"  end try\n" +
                                $"end tell";
            
            try
            {
                // Debug the AppleScript for troubleshooting
                AnsiConsole.MarkupLine($"[grey]Executing AppleScript for file selection[/]");
                
                // Create a temporary script file
                string scriptFile = Path.Combine(Path.GetTempPath(), $"applescript_{Guid.NewGuid()}.scpt");
                File.WriteAllText(scriptFile, appleScript);
                
                // Execute the AppleScript from the file
                var process = new Process
                {
                    StartInfo = new ProcessStartInfo
                    {
                        FileName = "osascript",
                        Arguments = scriptFile,
                        UseShellExecute = false,
                        RedirectStandardOutput = true,
                        RedirectStandardError = true,
                        CreateNoWindow = true
                    }
                };
                
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                string error = process.StandardError.ReadToEnd();
                process.WaitForExit();
                
                // Clean up the script file
                if (File.Exists(scriptFile))
                {
                    try { File.Delete(scriptFile); } catch { /* Ignore cleanup errors */ }
                }
                
                // If there's an error with AppleScript, log it
                if (!string.IsNullOrEmpty(error))
                {
                    AnsiConsole.MarkupLine($"[red]AppleScript error: {error}[/]");
                }
                
                // Read the selected path from the temporary file (if the dialog wasn't canceled)
                if (File.Exists(tempFile))
                {
                    string selectedPath = File.ReadAllText(tempFile).Trim();
                    File.Delete(tempFile); // Clean up temp file
                    return string.IsNullOrEmpty(selectedPath) ? null : selectedPath;
                }
                
                return null;
            }
            catch (Exception ex)
            {
                if (File.Exists(tempFile))
                {
                    try { File.Delete(tempFile); } catch { /* Ignore cleanup errors */ }
                }
                throw new Exception($"Failed to use macOS file dialog: {ex.Message}", ex);
            }
        }
        
        #endregion
    }
} 