using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using Independentsoft.Pst;
using Microsoft.Win32;
using System.IO;
using System.Windows.Forms; 
using PstMessage = Independentsoft.Pst.Message;
using System.Diagnostics;
using System.Threading;

namespace epasts
{
    public partial class MainWindow : Window
    {
        private List<string> pstFilePaths = new List<string>();
        private string outputDirectory = "";
        private int totalMessages = 0;
        private int processedMessages = 0;
        private const int RecipientThreshold = 20;
        private Stopwatch stopwatch;
        private int currentPstIndex = 0;
        private string currentPstFileName = "";
        private int totalAllMessages = 0;


        public MainWindow()
        {
            InitializeComponent();
        }

        private void SelectPstFiles_Click(object sender, RoutedEventArgs e)
        {
            var dialog = new Microsoft.Win32.OpenFileDialog
            {
                Filter = "Outlook PST files (*.pst)|*.pst",
                Multiselect = true,
                Title = "Choos PST files"
            };

            if (dialog.ShowDialog() == true)
            {
                pstFilePaths = dialog.FileNames.ToList();
                System.Windows.MessageBox.Show($"In total {pstFilePaths.Count} PST files have been chosen.");
            }
        }

        private void SelectOutputFolder_Click(object sender, RoutedEventArgs e)
        {
            using (var folderDialog = new FolderBrowserDialog())
            {
                if (folderDialog.ShowDialog() == System.Windows.Forms.DialogResult.OK)
                {
                    outputDirectory = folderDialog.SelectedPath;
                    System.Windows.MessageBox.Show($"Chosen folder: {outputDirectory}");
                }
            }
        }

        private async void StartProcessing_Click(object sender, RoutedEventArgs e)
        {
            if (pstFilePaths.Count == 0)
            {
                System.Windows.MessageBox.Show("Please first choos PST files.");
                return;
            }

            if (string.IsNullOrWhiteSpace(outputDirectory))
            {
                System.Windows.MessageBox.Show("Choose folder where to save .eml files.");
                return;
            }

            ProcessingProgressBar.Value = 0;
            StatusTextBlock.Text = "Calculates the count of emails...";

            // First, count all eligible messages
            // First, count all messages (for progress bar)
            StatusTextBlock.Text = "Calculates the count of emails...";
            totalAllMessages = await System.Threading.Tasks.Task.Run(() => CountAllMessagesInPstFiles());

            // Then count only eligible messages
            //StatusTextBlock.Text = "Skaita epastus ar daudz adresātiem...";
            //totalMessages = await System.Threading.Tasks.Task.Run(() => CountAllEligibleMessagesInPstFiles());
            totalMessages = totalAllMessages;


            processedMessages = 0;

            if (totalMessages == 0)
            {
                System.Windows.MessageBox.Show("There are no emails with the set criteria.");
                return;
            }

            ProcessingProgressBar.Maximum = totalAllMessages;
            stopwatch = Stopwatch.StartNew();

            // Then, process them
            await System.Threading.Tasks.Task.Run(() =>
            {
                currentPstIndex = 0;
                foreach (string pstFile in pstFilePaths)
                {
                    currentPstFileName = System.IO.Path.GetFileName(pstFile);
                    int localIndex = currentPstIndex + 1; // 1-based for display

                    Dispatcher.Invoke(() =>
                    {
                        StatusTextBlock.Text = $"Scanning {currentPstFileName} ({localIndex}/{pstFilePaths.Count})";
                    });
                    try
                    {
                        using (PstFile pst = new PstFile(pstFile))
                        {
                            // Start processing from the root of the PST file
                            // The ProcessFolder method will handle recursion
                            ProcessFolder(pst.Root, outputDirectory);
                        }
                    }
                    catch (Exception ex)
                    {
                        // Log the error instead of just showing a message box
                        // For a large file, frequent pop-ups are disruptive
                        // Consider using a logging framework like NLog or Serilog, or simply writing to a file.
                        // For demonstration, we'll use Console.WriteLine and Dispatcher for a single final message.
                        Console.WriteLine($"Error processing PST file {pstFile}: {ex.Message}");
                        Dispatcher.Invoke(() => System.Windows.MessageBox.Show($"Error processing {pstFile}:\n{ex.Message}\nCheck console for details."));
                    }
                    currentPstIndex++;
                }
            });

            stopwatch.Stop();
            ProcessingProgressBar.Value = totalMessages;
            // StatusTextBlock.Text = $"Pabeigts! Apstrādāti {totalMessages} epasti {stopwatch.Elapsed.TotalSeconds:F1} sekundēs.";
            Dispatcher.Invoke(() =>
            {
                ProcessingProgressBar.Value = processedMessages;
                double percent = (double)processedMessages / totalMessages * 100;

                TimeSpan elapsed = stopwatch.Elapsed;
                double avgPerMessage = processedMessages > 0 ? elapsed.TotalSeconds / processedMessages : 0;
                double remainingSeconds = avgPerMessage * (totalMessages - processedMessages);
                TimeSpan remainingTime = TimeSpan.FromSeconds(remainingSeconds);

                string currentFile = currentPstIndex < pstFilePaths.Count ? System.IO.Path.GetFileName(pstFilePaths[currentPstIndex]) : "";

                StatusTextBlock.Text =
                    $"Processed {processedMessages}/{totalAllMessages} e-mails ({percent:F1}%)" +
                    $"\nRemaining time: ~{remainingTime.Minutes:D2}m {remainingTime.Seconds:D2}s" +
                    $"\nFiles: {currentFile} ({currentPstIndex + 1}/{pstFilePaths.Count})";
            });
            System.Windows.MessageBox.Show("Processing of emails have finished!");
        }

        

        private int CountAllMessagesInPstFiles()
        {
            int total = 0;
            foreach (string pstFile in pstFilePaths)
            {
                try
                {
                    using (PstFile pst = new PstFile(pstFile))
                    {
                        total += CountAllMessages(pst.Root);
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Mistake while counting emails in the file {pstFile}: {ex.Message}");
                }
            }
            return total;
        }

        private int CountAllMessages(Folder folder)
        {
            int count = 0;

            if (folder.ContainerClass == "IPF.Note")
            {
                try
                {
                    count += folder.GetItems().Count;
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Mistake while accessing folder's '{folder.DisplayName}' emails: {ex.Message}");
                }
            }

            foreach (var subfolder in folder.GetFolders())
            {
                count += CountAllMessages(subfolder);
            }

            return count;
        }

        
        private void ProcessFolder(Folder folder, string outputDirectory)
        {
            if (folder.ContainerClass == "IPF.Note")
            {
                try
                {
                    foreach (Item item in folder.GetItems())
                    {
                        try
                        {
                            var message = item.GetMessageFile().ConvertToMimeMessage();
                            int toCount = message.To.Count;
                            int ccCount = message.Cc.Count;
                            int totalRecipients = toCount + ccCount;

                            if (totalRecipients >= RecipientThreshold)
                            {
                                string safeSubject = string.Join("_", message.Subject.Split(System.IO.Path.GetInvalidFileNameChars()));
                                if (string.IsNullOrWhiteSpace(safeSubject)) safeSubject = "NoSubject";

                                // Ensure unique filenames, especially when multiple threads/processes might be writing
                                // Adding a Guid might be safer for truly unique names, but DateTime.Now.Ticks is usually sufficient
                                string fileName = $"{DateTime.Now.Ticks}_{safeSubject}.eml"; // Using Ticks for higher uniqueness
                                string fullPath = System.IO.Path.Combine(outputDirectory, fileName);

                                message.Save(fullPath);
                            }

                            // Use Interlocked for thread-safe increment in async/multi-threaded scenarios
                            Interlocked.Increment(ref processedMessages);

                            // Update UI less frequently to improve performance.
                            // For example, update every 100 messages or on a timer.
                            // For simplicity, keeping it per-message for now, but be aware of performance.
                            Dispatcher.Invoke(() =>
                            {
                                ProcessingProgressBar.Value = processedMessages;
                                double percent = (double)processedMessages / totalAllMessages * 100;

                                TimeSpan elapsed = stopwatch.Elapsed;
                                double avgPerMessage = processedMessages > 0 ? elapsed.TotalSeconds / processedMessages : 0;
                                double remainingSeconds = avgPerMessage * (totalAllMessages - processedMessages);
                                TimeSpan remainingTime = TimeSpan.FromSeconds(remainingSeconds);

                                StatusTextBlock.Text =
                                    $"Processed {processedMessages}/{totalAllMessages} emails ({percent:F1}%)" +
                                    $"\nRemaining time: ~{remainingTime.Minutes:D2}m {remainingTime.Seconds:D2}s" +
                                    $"\nActive file: {currentPstFileName} ({currentPstIndex + 1}/{pstFilePaths.Count})";
                            });
                        }
                        catch (Exception ex)
                        {
                            // Log error for individual message
                            Console.WriteLine($"Error processing message in folder '{folder.DisplayName}': {ex.Message}");
                        }
                    }
                }
                catch (Exception ex)
                {
                    // Log error for folder item enumeration
                    Console.WriteLine($"Error enumerating items in folder '{folder.DisplayName}': {ex.Message}");
                }
            }

            // Crucially, recurse into all subfolders
            foreach (Folder subfolder in folder.GetFolders())
            {
                ProcessFolder(subfolder, outputDirectory);
            }
        }
        
    }
}