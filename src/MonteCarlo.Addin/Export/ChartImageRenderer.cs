using System.IO;
using System.Windows;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using SkiaSharp;

namespace MonteCarlo.Addin.Export;

/// <summary>
/// Renders WPF and SkiaSharp chart controls to PNG byte arrays for embedding in Excel.
/// </summary>
public static class ChartImageRenderer
{
    /// <summary>
    /// Render a WPF FrameworkElement to a PNG byte array.
    /// </summary>
    /// <param name="control">The WPF control to render.</param>
    /// <param name="width">Desired pixel width.</param>
    /// <param name="height">Desired pixel height.</param>
    /// <param name="dpi">Output DPI (default 192 for high-DPI).</param>
    public static byte[] RenderWpfControl(FrameworkElement control, int width, int height, double dpi = 192)
    {
        control.Measure(new Size(width, height));
        control.Arrange(new Rect(0, 0, width, height));
        control.UpdateLayout();

        int pixelWidth = (int)(width * dpi / 96);
        int pixelHeight = (int)(height * dpi / 96);

        var renderBitmap = new RenderTargetBitmap(pixelWidth, pixelHeight, dpi, dpi, PixelFormats.Pbgra32);
        renderBitmap.Render(control);

        var encoder = new PngBitmapEncoder();
        encoder.Frames.Add(BitmapFrame.Create(renderBitmap));

        using var stream = new MemoryStream();
        encoder.Save(stream);
        return stream.ToArray();
    }

    /// <summary>
    /// Render a SkiaSharp drawing action to a PNG byte array.
    /// Used for the tornado chart and other SkiaSharp-based controls.
    /// </summary>
    /// <param name="width">Desired pixel width.</param>
    /// <param name="height">Desired pixel height.</param>
    /// <param name="drawAction">Action that draws onto the SKCanvas.</param>
    /// <param name="scale">Scale factor for high-DPI (default 2).</param>
    public static byte[] RenderSkiaSharp(int width, int height, Action<SKCanvas, SKImageInfo> drawAction, int scale = 2)
    {
        int scaledWidth = width * scale;
        int scaledHeight = height * scale;

        using var bitmap = new SKBitmap(scaledWidth, scaledHeight);
        using var canvas = new SKCanvas(bitmap);
        canvas.Clear(SKColors.White);
        canvas.Scale(scale);

        var info = new SKImageInfo(scaledWidth, scaledHeight);
        drawAction(canvas, info);

        using var image = SKImage.FromBitmap(bitmap);
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        return data.ToArray();
    }

    /// <summary>
    /// Save PNG bytes to a temporary file. Returns the temp file path.
    /// Caller is responsible for cleanup.
    /// </summary>
    public static string SaveToTempFile(byte[] pngBytes)
    {
        string tempPath = Path.Combine(Path.GetTempPath(), $"mc_chart_{Guid.NewGuid():N}.png");
        File.WriteAllBytes(tempPath, pngBytes);
        return tempPath;
    }
}
