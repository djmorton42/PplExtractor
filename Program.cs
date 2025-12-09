using Avalonia;
using System;

namespace PplExtractor;

// Register code page provider for ExcelDataReader
sealed class Program
{
    [STAThread]
    public static void Main(string[] args)
    {
        // Register code page provider for ExcelDataReader
        System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
        
        BuildAvaloniaApp()
            .StartWithClassicDesktopLifetime(args);
    }

    // Avalonia configuration, don't remove; also used by visual designer.
    public static AppBuilder BuildAvaloniaApp()
        => AppBuilder.Configure<App>()
            .UsePlatformDetect()
            .WithInterFont()
            .LogToTrace();
}
