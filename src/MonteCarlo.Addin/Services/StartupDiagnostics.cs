using System.Reflection;
using System.IO;
using ExcelDna.Integration;

namespace MonteCarlo.Addin.Services;

/// <summary>
/// Writes lightweight startup and runtime diagnostics for Excel loading issues.
/// </summary>
internal static class StartupDiagnostics
{
    private static readonly object Lock = new();

    public static string LogDirectory =>
        Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "MonteCarlo.XL",
            "Logs");

    public static string LogPath => Path.Combine(LogDirectory, "startup.log");

    public static void LogStartup()
    {
        try
        {
            var app = (Application)ExcelDnaUtil.Application;
            Log(
                $"Startup: add-in version={GetVersion()}, process={(Environment.Is64BitProcess ? "x64" : "x86")}, " +
                $"ExcelVersion={app.Version}, ExcelBuild={app.Build}, addInPath={ExcelDnaUtil.XllPath}");
        }
        catch (Exception ex)
        {
            LogException("Startup diagnostics failed.", ex);
        }
    }

    public static void Log(string message)
    {
        try
        {
            lock (Lock)
            {
                Directory.CreateDirectory(LogDirectory);
                File.AppendAllText(LogPath, $"[{DateTimeOffset.Now:O}] {message}{Environment.NewLine}");
            }
        }
        catch
        {
            // Diagnostics must never prevent the add-in from loading.
        }
    }

    public static void LogException(string message, Exception exception)
    {
        Log($"{message} {exception.GetType().Name}: {exception.Message}{Environment.NewLine}{exception}");
    }

    private static string GetVersion()
    {
        return Assembly.GetExecutingAssembly().GetName().Version?.ToString() ?? "unknown";
    }
}
