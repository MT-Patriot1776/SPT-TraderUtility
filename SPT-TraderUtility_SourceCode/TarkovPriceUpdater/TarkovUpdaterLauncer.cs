
// C:\TarkovPriceUpdater\TarkovUpdaterLauncher.cs

// --- Using directives (must be first for legacy compiler) ---
using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using System.Runtime.InteropServices;

// --- Assembly metadata (appears in EXE Properties -> Details) ---
[assembly: AssemblyTitle("Tarkov Trader Price Updater")]
[assembly: AssemblyDescription("Launcher that prefers PowerShell 7 (pwsh) with Windows PowerShell fallback.")]
[assembly: AssemblyCompany("MT_Militia")]
[assembly: AssemblyProduct("Tarkov Trader Price Updater")]
[assembly: AssemblyCopyright("© 2025 MT_Militia")]
[assembly: AssemblyVersion("1.0.0.0")]
[assembly: AssemblyFileVersion("1.0.0.0")]
[assembly: ComVisible(false)]

namespace TarkovUpdaterLauncher
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            string tempScript = null;

            try
            {
                // 1) Try embedded script first
                Stream res = Assembly.GetExecutingAssembly()
                                     .GetManifestResourceStream("TarkovPriceUpdater.ps1");

                if (res != null)
                {
                    tempScript = Path.Combine(Path.GetTempPath(),
                        "TarkovUpdater_" + Guid.NewGuid().ToString("N") + ".ps1");
                    using (FileStream fs = new FileStream(tempScript, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        res.CopyTo(fs);
                    }
                }
                else
                {
                    // 2) Fallback to external script on disk
                    string externalScript = @"C:\TarkovPriceUpdater\TarkovPriceUpdater.ps1";
                    if (File.Exists(externalScript))
                    {
                        tempScript = externalScript;
                    }
                    else
                    {
                        MessageBox.Show(
                            "Embedded script not found and external script missing:\n" + externalScript +
                            "\n\nIf you want a single EXE, recompile with:\n/resource:TarkovPriceUpdater.ps1,TarkovPriceUpdater.ps1",
                            "Tarkov Trader Price Updater",
                            MessageBoxButtons.OK, MessageBoxIcon.Error);
                        return;
                    }
                }

                // 3) Choose PS host: prefer pwsh.exe (PS7), fallback to powershell.exe (PS5.1)
                string hostExe = FindPwsh();
                if (string.IsNullOrWhiteSpace(hostExe))
                    hostExe = GetWinPSPath();

                string hostArgs = "-NoProfile -ExecutionPolicy Bypass -STA -File \"" + tempScript + "\"";

                // 4) Launch hidden (no console)
                ProcessStartInfo psi = new ProcessStartInfo();
                psi.FileName = hostExe;
                psi.Arguments = hostArgs;
                psi.UseShellExecute = false;
                psi.CreateNoWindow = true;
                psi.WindowStyle = ProcessWindowStyle.Hidden;
                psi.WorkingDirectory = Path.GetDirectoryName(tempScript);

                using (Process p = Process.Start(psi))
                {
                    if (p != null) p.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to launch updater:\n" + ex.Message,
                    "Tarkov Trader Price Updater", MessageBoxButtons.OK, MessageBoxIcon.Error);
            }
            finally
            {
                // Clean up temp file if we extracted the embedded script
                try
                {
                    if (!string.IsNullOrWhiteSpace(tempScript)
                        && tempScript.StartsWith(Path.GetTempPath(), StringComparison.OrdinalIgnoreCase)
                        && File.Exists(tempScript))
                    {
                        File.Delete(tempScript);
                    }
                }
                catch { /* ignore cleanup errors */ }
            }
        }

        // Prefer PS7 (pwsh.exe)
        private static string FindPwsh()
        {
            string[] candidates = new string[]
            {
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "PowerShell", "7", "pwsh.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles), "PowerShell", "7-preview", "pwsh.exe"),
                Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "Microsoft", "WindowsApps", "pwsh.exe")
            };

            foreach (string c in candidates)
            {
                try { if (File.Exists(c)) return c; } catch { }
            }

            string path = Environment.GetEnvironmentVariable("PATH");
            if (!string.IsNullOrEmpty(path))
            {
                foreach (string dir in path.Split(Path.PathSeparator))
                {
                    try
                    {
                        if (string.IsNullOrWhiteSpace(dir)) continue;
                        string test = Path.Combine(dir.Trim(), "pwsh.exe");
                        if (File.Exists(test)) return test;
                    }
                    catch { }
                }
            }
            return null;
        }

        // Fallback to Windows PowerShell 5.1
        private static string GetWinPSPath()
        {
            string winDir = Environment.GetFolderPath(Environment.SpecialFolder.Windows);
            string p1 = Path.Combine(winDir, "System32", "WindowsPowerShell", "v1.0", "powershell.exe");
            string p2 = Path.Combine(winDir, "SysWOW64", "WindowsPowerShell", "v1.0", "powershell.exe");
            if (File.Exists(p1)) return p1;
            if (File.Exists(p2)) return p2;
            return "powershell.exe"; // last resort
        }
    }
}
