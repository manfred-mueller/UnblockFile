using System;
using System.Diagnostics;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace UnblockFile
{
    class Program
    {
        static async Task Main(string[] args)
        {
            if (args.Length > 0)
            {
                string action = args[0].ToLower();
                if (action == "register")
                {
                    RegisterContextMenu();
                }
                else if (action == "unregister")
                {
                    UnregisterContextMenu();
                }
                else
                {
                    string path = args[0];
                    if (File.Exists(path))
                    {
                        await UnblockFile(path);
                    }
                    else if (Directory.Exists(path))
                    {
                        await UnblockFilesInDirectory(path);
                    }
                    else
                    {
                        ShowBalloonTip(Properties.Resources.Error, Properties.Resources.FileOrFolderNotFound, ToolTipIcon.Error);
                    }
                }
            }
            else
            {
                ShowBalloonTip("", Properties.Resources.UsageTunblockFile_pathNTunblockRegisterNTunblockUnregister, ToolTipIcon.Warning);
                return;
            }
        }

        static async Task UnblockFile(string filePath)
        {
            try
            {
                // Attempt to unblock the file
                bool unblocked = await TryUnblockAsync(filePath);

                if (unblocked)
                {
                    ShowBalloonTip(Properties.Resources.FileUnblocked, Properties.Resources.TheFileHasBeenSuccessfullyUnblocked, ToolTipIcon.Info);
                }
                else
                {
                    ShowBalloonTip(Properties.Resources.Error, Properties.Resources.FailedToUnblockTheFile, ToolTipIcon.Error);
                }
            }
            catch (Exception ex)
            {
                ShowBalloonTip(Properties.Resources.Error, Properties.Resources.AnErrorOccurred + ex.Message, ToolTipIcon.Error);
            }
        }

        static async Task UnblockFilesInDirectory(string directoryPath)
        {
            int totalFiles = Directory.GetFiles(directoryPath, "*", SearchOption.AllDirectories).Length;
            int processedFiles = 0;

            foreach (string filePath in Directory.GetFiles(directoryPath, "*", SearchOption.AllDirectories))
            {
                try
                {
                    bool unblocked = await TryUnblockAsync(filePath);
                    if (unblocked)
                    {
                        processedFiles++;
                        if (processedFiles == totalFiles)
                        {
                            ShowBalloonTip(Properties.Resources.FileUnblocked, Properties.Resources.TheFileHasBeenSuccessfullyUnblocked, ToolTipIcon.Info);
                        }
                    }
                    else
                    {
                        ShowBalloonTip(Properties.Resources.Error, Properties.Resources.FailedToUnblockTheFile, ToolTipIcon.Error);
                    }
                }
                catch (Exception)
                {
                    ShowBalloonTip(Properties.Resources.Error, Properties.Resources.FailedToUnblockTheFile, ToolTipIcon.Error);
                }
            }
        }

        static async Task<bool> TryUnblockAsync(string filePath)
        {
            try
            {
                // Check if the file exists
                if (File.Exists(filePath))
                {
                    // Remove Zone.Identifier alternate data stream
                    await RemoveZoneIdentifierAsync(filePath);
                    return true;
                }
                else
                {
                    ShowBalloonTip(Properties.Resources.Error, Properties.Resources.FileOrFolderNotFound, ToolTipIcon.Error);
                    return false;
                }
            }
            catch (UnauthorizedAccessException)
            {
                // Insufficient permissions to unblock the file
                return false;
            }
        }

        static async Task RemoveZoneIdentifierAsync(string filePath)
        {
            // Check if the file has Zone.Identifier alternate data stream
            if (File.Exists(filePath + ":Zone.Identifier"))
            {
                // Delete the Zone.Identifier alternate data stream
                await Task.Run(() => File.Delete(filePath + ":Zone.Identifier"));
            }
        }

        static void ShowBalloonTip(string title, string message, ToolTipIcon icon)
        {
            NotifyIcon notifyIcon = new NotifyIcon();
            notifyIcon.Visible = true;

            // Load icon from resources
            notifyIcon.Icon = Properties.Resources.UnblockFileIcon;

            notifyIcon.BalloonTipIcon = icon;
            notifyIcon.BalloonTipTitle = title;
            notifyIcon.BalloonTipText = message;
            notifyIcon.ShowBalloonTip(3000);
            notifyIcon.Dispose();
        }

        static void RegisterContextMenu()
        {
            string exePath = Process.GetCurrentProcess().MainModule.FileName;
            string regTextFile = Properties.Resources.UnblockFile;
            string regTextFolder = Properties.Resources.UnblockFolder;
            string regCommand = $"\"{exePath}\" \"%V\"";
            string iconPath = $"\"{exePath}\",0"; // Path to the executable and icon index (0 by default)

            // For files
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\Classes\*\shell\UnblockFile"))
            {
                key.SetValue("", regTextFile);
                key.CreateSubKey("command").SetValue("", regCommand);
                key.SetValue("Icon", iconPath, RegistryValueKind.String);
            }

            // For directories
            using (RegistryKey key = Registry.CurrentUser.CreateSubKey(@"Software\Classes\Directory\shell\UnblockFile"))
            {
                key.SetValue("", regTextFolder);
                key.CreateSubKey("command").SetValue("", regCommand);
                key.SetValue("Icon", iconPath, RegistryValueKind.String);
            }

            ShowBalloonTip("", Properties.Resources.ProgramRegisteredInTheContextMenuSuccessfully, ToolTipIcon.Info);
        }

        static void UnregisterContextMenu()
        {
            Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\*\shell\UnblockFile", false);
            Registry.CurrentUser.DeleteSubKeyTree(@"Software\Classes\Directory\shell\UnblockFile", false);
            ShowBalloonTip("", Properties.Resources.ProgramUnregisteredFromTheContextMenuSuccessfully, ToolTipIcon.Info);
        }
    }
}
