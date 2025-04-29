using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Spectre.Console;

namespace InvoiceMailerUI
{
    /// <summary>
    /// A simple console-based file and folder browser using Spectre.Console
    /// </summary>
    public class FileBrowser
    {
        private const string PARENT_DIRECTORY = ".. (Parent Directory)";
        private const string CURRENT_DIRECTORY = ". (Select Current Directory)";
        private const string CREATE_DIRECTORY = "Create New Directory";
        
        /// <summary>
        /// Browse for a folder
        /// </summary>
        /// <param name="startPath">The starting path for browsing</param>
        /// <returns>The selected folder path or null if canceled</returns>
        public static string? BrowseForFolder(string? startPath = null)
        {
            // Default to current directory if startPath is null or empty
            string currentPath = !string.IsNullOrEmpty(startPath) && Directory.Exists(startPath)
                ? startPath
                : Directory.GetCurrentDirectory();
            
            while (true)
            {
                try
                {
                    AnsiConsole.Clear();
                    AnsiConsole.Write(new Rule($"[yellow]Folder Browser: {currentPath}[/]").RuleStyle("grey").Centered());
                    AnsiConsole.WriteLine();
                    
                    // Get directories in the current path
                    var directories = Directory.GetDirectories(currentPath)
                        .Select(dir => Path.GetFileName(dir))
                        .Where(name => name != null) // Filter out null values
                        .Cast<string>() // Ensure all items are strings
                        .OrderBy(d => d)
                        .ToList();
                    
                    var choices = new List<string>
                    {
                        CURRENT_DIRECTORY,
                        PARENT_DIRECTORY,
                        CREATE_DIRECTORY
                    };
                    
                    choices.AddRange(directories);
                    
                    var choice = AnsiConsole.Prompt(
                        new SelectionPrompt<string>()
                            .Title("Select a folder or navigate:")
                            .PageSize(20)
                            .AddChoices(choices));
                    
                    if (choice == CURRENT_DIRECTORY)
                    {
                        return currentPath;
                    }
                    else if (choice == PARENT_DIRECTORY)
                    {
                        var parent = Directory.GetParent(currentPath);
                        if (parent != null)
                        {
                            currentPath = parent.FullName;
                        }
                    }
                    else if (choice == CREATE_DIRECTORY)
                    {
                        var dirName = AnsiConsole.Prompt(
                            new TextPrompt<string>("Enter new directory name:")
                                .Validate(name =>
                                {
                                    if (string.IsNullOrWhiteSpace(name))
                                        return ValidationResult.Error("Directory name cannot be empty");
                                    if (name.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                                        return ValidationResult.Error("Directory name contains invalid characters");
                                    if (Directory.Exists(Path.Combine(currentPath, name)))
                                        return ValidationResult.Error("Directory already exists");
                                    return ValidationResult.Success();
                                }));
                        
                        try
                        {
                            var newDirPath = Path.Combine(currentPath, dirName);
                            Directory.CreateDirectory(newDirPath);
                            currentPath = newDirPath;
                        }
                        catch (Exception ex)
                        {
                            AnsiConsole.MarkupLine($"[red]Error creating directory: {ex.Message}[/]");
                            AnsiConsole.WriteLine();
                            AnsiConsole.WriteLine("Press any key to continue...");
                            Console.ReadKey();
                        }
                    }
                    else
                    {
                        currentPath = Path.Combine(currentPath, choice);
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    AnsiConsole.MarkupLine("[red]Access denied to this directory. Please select another path.[/]");
                    var parent = Directory.GetParent(currentPath);
                    if (parent != null)
                    {
                        currentPath = parent.FullName;
                    }
                    AnsiConsole.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                }
                catch (Exception ex)
                {
                    AnsiConsole.MarkupLine($"[red]Error: {ex.Message}[/]");
                    AnsiConsole.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                    return null;
                }
            }
        }
        
        /// <summary>
        /// Browse for a file
        /// </summary>
        /// <param name="startPath">The starting path for browsing</param>
        /// <param name="filter">The file filter (e.g., "*.csv")</param>
        /// <returns>The selected file path or null if canceled</returns>
        public static string? BrowseForFile(string? startPath = null, string filter = "*.*")
        {
            // Default to current directory if startPath is null or empty
            string currentPath = !string.IsNullOrEmpty(startPath) && Directory.Exists(startPath)
                ? startPath
                : Directory.GetCurrentDirectory();
            
            while (true)
            {
                try
                {
                    AnsiConsole.Clear();
                    AnsiConsole.Write(new Rule($"[yellow]File Browser: {currentPath}[/]").RuleStyle("grey").Centered());
                    AnsiConsole.WriteLine();
                    
                    // Get directories in the current path
                    var directories = Directory.GetDirectories(currentPath)
                        .Select(d => $"[blue]{Path.GetFileName(d)}/[/]")
                        .Where(name => name != null) // Filter out null values
                        .Cast<string>() // Ensure all items are strings
                        .OrderBy(d => d)
                        .ToList();
                    
                    // Get files matching the filter in the current path
                    var files = Directory.GetFiles(currentPath, filter)
                        .Select(f => $"[green]{Path.GetFileName(f)}[/]")
                        .Where(name => name != null) // Filter out null values
                        .Cast<string>() // Ensure all items are strings
                        .OrderBy(f => f)
                        .ToList();
                    
                    var choices = new List<string>
                    {
                        PARENT_DIRECTORY,
                        "Cancel"
                    };
                    
                    choices.AddRange(directories);
                    choices.AddRange(files);
                    
                    var choice = AnsiConsole.Prompt(
                        new SelectionPrompt<string>()
                            .Title($"Select a file or navigate (Filter: {filter}):")
                            .PageSize(20)
                            .AddChoices(choices));
                    
                    if (choice == "Cancel")
                    {
                        return null;
                    }
                    else if (choice == PARENT_DIRECTORY)
                    {
                        var parent = Directory.GetParent(currentPath);
                        if (parent != null)
                        {
                            currentPath = parent.FullName;
                        }
                    }
                    else
                    {
                        // Strip markup for processing
                        string cleanChoice = choice.Replace("[blue]", "").Replace("[/]", "").Replace("[green]", "").Replace("/", "");
                        
                        string selectedPath = Path.Combine(currentPath, cleanChoice);
                        
                        if (Directory.Exists(selectedPath))
                        {
                            currentPath = selectedPath;
                        }
                        else if (File.Exists(selectedPath))
                        {
                            return selectedPath;
                        }
                    }
                }
                catch (UnauthorizedAccessException)
                {
                    AnsiConsole.MarkupLine("[red]Access denied to this directory. Please select another path.[/]");
                    var parent = Directory.GetParent(currentPath);
                    if (parent != null)
                    {
                        currentPath = parent.FullName;
                    }
                    AnsiConsole.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                }
                catch (Exception ex)
                {
                    AnsiConsole.MarkupLine($"[red]Error: {ex.Message}[/]");
                    AnsiConsole.WriteLine("Press any key to continue...");
                    Console.ReadKey();
                    return null;
                }
            }
        }
    }
} 