
// TarkovAssortBuilderLauncher.cs
using System;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;

namespace TarkovAssortBuilderLauncher
{
    internal static class Program
    {
        [STAThread]
        private static void Main(string[] args)
        {
            string tempScript = null;
            try
            {
                // 1) Extract embedded .ps1 resource to %TEMP%
                var asm = Assembly.GetExecutingAssembly();
                var resName = asm.GetManifestResourceNames()
                                 .FirstOrDefault(n => n.EndsWith(".ps1", StringComparison.OrdinalIgnoreCase));
                if (resName != null)
                {
                    tempScript = Path.Combine(Path.GetTempPath(),
                        "TarkovAssort_" + Guid.NewGuid().ToString("N") + ".ps1");
                    using (var res = asm.GetManifestResourceStream(resName))
                    using (var fs = new FileStream(tempScript, FileMode.Create, FileAccess.Write, FileShare.None))
                    {
                        res!.CopyTo(fs);
                    }
                }
                else
                {
                    MessageBox.Show(
                        "Embedded script not found. Recompile with the .ps1 as EmbeddedResource.",
                        "Tarkov Assort Builder",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    return;
                }

                // 2) Use native Windows PowerShell so the embedded script runs on a stock Win11 install.
                string hostExe = GetWinPSPath();

                string hostArgs = "-NoProfile -ExecutionPolicy Bypass -STA -WindowStyle Hidden -File \"" + tempScript + "\"";

                // 3) Let PowerShell hide only its console; the WinForms script window should still show.
                var psi = new ProcessStartInfo
                {
                    FileName = hostExe,
                    Arguments = hostArgs,
                    UseShellExecute = false,
                    CreateNoWindow = false,
                    WindowStyle = ProcessWindowStyle.Normal,
                    WorkingDirectory = Path.GetDirectoryName(tempScript)!
                };

                using (var p = Process.Start(psi))
                {
                    if (p != null) p.WaitForExit();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("Failed to launch Assort Builder:\n" + ex.Message,
                    "Tarkov Assort Builder",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
            }
            finally
            {
                // 4) Clean up temp file
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

        // Windows PowerShell 5.1 ships with Windows 11.
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
