using System;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows.Forms;
using Microsoft.Win32;

namespace Unblock
{
    class Program
    {
        // Import the GetFileAttributes function from kernel32.dll
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern uint GetFileAttributes(string lpFileName);

        // Import the DeleteFile function from kernel32.dll
        [DllImport("kernel32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern bool DeleteFile(string path);

        // Constants for GetFileAttributes
        private const uint INVALID_FILE_ATTRIBUTES = 0xFFFFFFFF;
        private const uint FILE_ATTRIBUTE_NORMAL = 0x80;

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

        static async Task UnblockFilesInDirectory(string directoryPath)
        {
            int totalFiles = Directory.GetFiles(directoryPath, "*", SearchOption.AllDirectories).Length;
            int processedFiles = 0;

            await Task.Run(() =>
            {
                RecursivelyUnblock(directoryPath, ref processedFiles, totalFiles);
            });

            if (processedFiles == totalFiles)
            {
                ShowBalloonTip(Properties.Resources.FileUnblocked, Properties.Resources.AllFilesHaveBeenSuccessfullyUnblocked, ToolTipIcon.Info);
            }
        }

        static void RecursivelyUnblock(string directoryPath, ref int processedFiles, int totalFiles)
        {
            foreach (string filePath in Directory.GetFiles(directoryPath, "*", SearchOption.TopDirectoryOnly))
            {
                try
                {
                    bool unblocked = TryUnblock(filePath);
                    if (unblocked)
                    {
                        processedFiles++;
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

            foreach (string subdirectory in Directory.GetDirectories(directoryPath))
            {
                RecursivelyUnblock(subdirectory, ref processedFiles, totalFiles);
            }
        }

        static bool TryUnblock(string filePath)
        {
            try
            {
                if (File.Exists(filePath))
                {
                    if (IsFileBlocked(filePath))
                    {
                        UnblockFile(filePath).Wait(); // Synchronously wait for file unblocking
                        return true;
                    }
                    else
                    {
                        ShowBalloonTip(Properties.Resources.Error, Properties.Resources.FileIsNotBlocked, ToolTipIcon.Error);
                        return false;
                    }
                }
                else
                {
                    ShowBalloonTip(Properties.Resources.Error, Properties.Resources.FileOrFolderNotFound, ToolTipIcon.Error);
                    return false;
                }
            }
            catch (UnauthorizedAccessException)
            {
                ShowBalloonTip(Properties.Resources.Error, Properties.Resources.InsufficientPermissions, ToolTipIcon.Error);
                return false;
            }
        }

        static bool IsFileBlocked(string filePath)
        {
            string zoneIdentifierPath = filePath + ":Zone.Identifier";
            uint fileAttributes = GetFileAttributes(zoneIdentifierPath);

            return (fileAttributes != INVALID_FILE_ATTRIBUTES && (fileAttributes & FILE_ATTRIBUTE_NORMAL) == 0);
        }

        static async Task UnblockFile(string filePath)
        {
            if (!IsFileBlocked(filePath))
            {
                ShowBalloonTip(Properties.Resources.Error, Properties.Resources.FileIsNotBlocked, ToolTipIcon.Error);
                return;
            }

            await Task.Run(() =>
            {
                try
                {
                    bool deleteSuccess = DeleteFile(filePath + ":Zone.Identifier");
                    if (!deleteSuccess)
                    {
                        int errorCode = Marshal.GetLastWin32Error();
                        if (errorCode != 2) // ERROR_FILE_NOT_FOUND
                        {
                            throw new System.ComponentModel.Win32Exception(errorCode);
                        }
                    }
                }
                catch (Exception ex)
                {
                    throw new Exception($"Failed to delete Zone.Identifier for file {filePath}. Error: {ex.Message}");
                }
            });
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
