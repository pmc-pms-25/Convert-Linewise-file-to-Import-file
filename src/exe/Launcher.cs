using System;
using System.Diagnostics;
using System.IO;
using System.Reflection;
using System.Windows.Forms;
using Microsoft.Win32;

[assembly: AssemblyTitle("Convert MTO")]
[assembly: AssemblyProduct("Convert MTO - Piping / Structural / EIT")]
[assembly: AssemblyVersion("1.0.0.0")]

static class Program
{
    [STAThread]
    static void Main()
    {
        try
        {
            string dir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "ConvertMTO");
            Directory.CreateDirectory(dir);
            string html = Path.Combine(dir, "Convert MTO.html");
            Extract("xlsx.full.min.js", Path.Combine(dir, "xlsx.full.min.js"));
            Extract("app.html", html);

            string url = new Uri(html).AbsoluteUri;
            string browser = FindBrowser();
            if (browser == null)
            {
                Process.Start(html);
                return;
            }
            string profile = Path.Combine(dir, "Profile");
            string args = "--app=\"" + url + "\" --user-data-dir=\"" + profile + "\" --no-first-run --no-default-browser-check --window-size=1360,920";
            Process.Start(new ProcessStartInfo(browser, args) { UseShellExecute = false });
        }
        catch (Exception ex)
        {
            MessageBox.Show("Không mở được Convert MTO:\n" + ex.Message, "Convert MTO", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    // write the embedded resource next to the page, only when it changed (new exe version)
    static void Extract(string resource, string path)
    {
        byte[] data;
        using (Stream s = Assembly.GetExecutingAssembly().GetManifestResourceStream(resource))
        using (MemoryStream ms = new MemoryStream()) { s.CopyTo(ms); data = ms.ToArray(); }
        if (!File.Exists(path) || !SameBytes(File.ReadAllBytes(path), data)) File.WriteAllBytes(path, data);
    }

    static bool SameBytes(byte[] a, byte[] b)
    {
        if (a.Length != b.Length) return false;
        for (int i = 0; i < a.Length; i++) if (a[i] != b[i]) return false;
        return true;
    }

    static string FindBrowser()
    {
        string pf86 = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFilesX86);
        string pf = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles);
        string la = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        string[] candidates = {
            Path.Combine(pf86, @"Microsoft\Edge\Application\msedge.exe"),
            Path.Combine(pf, @"Microsoft\Edge\Application\msedge.exe"),
            AppPath("msedge.exe"),
            Path.Combine(pf, @"Google\Chrome\Application\chrome.exe"),
            Path.Combine(pf86, @"Google\Chrome\Application\chrome.exe"),
            Path.Combine(la, @"Google\Chrome\Application\chrome.exe"),
            AppPath("chrome.exe")
        };
        foreach (string c in candidates) if (!string.IsNullOrEmpty(c) && File.Exists(c)) return c;
        return null;
    }

    static string AppPath(string exe)
    {
        foreach (RegistryKey root in new[] { Registry.LocalMachine, Registry.CurrentUser })
        {
            using (RegistryKey k = root.OpenSubKey(@"SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\" + exe))
            {
                if (k != null) { object v = k.GetValue(""); if (v != null) return v.ToString().Trim('"'); }
            }
        }
        return null;
    }
}
