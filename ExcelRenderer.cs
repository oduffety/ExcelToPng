using ClosedXML.Excel;
using SkiaSharp;

namespace ExcelAgent.Services;

/// <summary>
/// Renders Excel worksheets to images using SkiaSharp
/// </summary>
public class ExcelRenderer
{
    private const int WorksheetNameHeight = 30;
    private const int ColumnHeaderHeight = 20;
    private const int RowHeaderWidth = 40;
    private const int MaxColumnsToRender = 20;
    private const int MaxRowsToRender = 50;
    private const double PixelsPerExcelWidth = 7.5; // Excel column width units to pixels
    private const double DefaultColumnWidth = 64.0; // Default column width in pixels
    private const double DefaultRowHeight = 20.0; // Default row height in pixels
    
    /// <summary>
    /// Render a worksheet to a PNG image
    /// </summary>
    public byte[] RenderWorksheetToPng(IXLWorksheet worksheet, int maxColumns = MaxColumnsToRender, int maxRows = MaxRowsToRender)
    {
        var usedRange = worksheet.RangeUsed();
        if (usedRange == null)
        {
            // Empty worksheet - create a small placeholder image
            return CreateEmptyWorksheetImage();
        }
        
        var firstRow = 1;
        var firstCol = 1;
        var lastRow = Math.Min(usedRange.LastRow().RowNumber(), firstRow + maxRows - 1);
        var lastCol = Math.Min(usedRange.LastColumn().ColumnNumber(), firstCol + maxColumns - 1);
        
        // Calculate actual column widths and row heights
        var columnWidths = new Dictionary<int, double>();
        var rowHeights = new Dictionary<int, double>();
        var columnPositions = new Dictionary<int, double>();
        var rowPositions = new Dictionary<int, double>();
        
        double totalWidth = RowHeaderWidth; // Start after row header
        for (int col = firstCol; col <= lastCol; col++)
        {
            var width = worksheet.Column(col).Width * PixelsPerExcelWidth;
            if (width < 1) width = DefaultColumnWidth; // Use default if width is 0 or very small
            columnWidths[col] = width;
            columnPositions[col] = totalWidth;
            totalWidth += width;
        }
        
        double totalHeight = WorksheetNameHeight + ColumnHeaderHeight; // Start after worksheet name and column headers
        for (int row = firstRow; row <= lastRow; row++)
        {
            var height = worksheet.Row(row).Height;
            if (height < 1) height = DefaultRowHeight; // Use default if height is 0 or very small
            rowHeights[row] = height;
            rowPositions[row] = totalHeight;
            totalHeight += height;
        }
        
        var imageWidth = (int)Math.Ceiling(totalWidth) + 1;
        var imageHeight = (int)Math.Ceiling(totalHeight) + 1;
        
        var imageInfo = new SKImageInfo(imageWidth, imageHeight);
        using var surface = SKSurface.Create(imageInfo);
        var canvas = surface.Canvas;
        
        // White background
        canvas.Clear(SKColors.White);
        
        // Draw worksheet name header
        DrawWorksheetHeader(canvas, worksheet.Name, imageWidth);
        
        // Draw column headers (A, B, C, etc.)
        DrawColumnHeaders(canvas, firstCol, lastCol, columnPositions, columnWidths);
        
        // Draw row headers (1, 2, 3, etc.)
        DrawRowHeaders(canvas, firstRow, lastRow, rowPositions, rowHeights);
        
        // Track which cells are already drawn (for merged cells)
        var drawnCells = new HashSet<string>();
        
        // Draw cells
        for (int row = firstRow; row <= lastRow; row++)
        {
            for (int col = firstCol; col <= lastCol; col++)
            {
                var cell = worksheet.Cell(row, col);
                var cellKey = $"{row},{col}";
                
                // Skip if already drawn as part of a merged cell
                if (drawnCells.Contains(cellKey))
                    continue;
                
                var x = columnPositions[col];
                var y = rowPositions[row];
                var width = columnWidths[col];
                var height = rowHeights[row];
                
                // Check if this is a merged cell
                if (cell.IsMerged())
                {
                    var mergedRange = cell.MergedRange();
                    var mergeFirstRow = mergedRange.FirstRow().RowNumber();
                    var mergeFirstCol = mergedRange.FirstColumn().ColumnNumber();
                    var mergeLastRow = mergedRange.LastRow().RowNumber();
                    var mergeLastCol = mergedRange.LastColumn().ColumnNumber();
                    
                    // Calculate merged cell dimensions
                    width = 0;
                    for (int c = mergeFirstCol; c <= mergeLastCol && c <= lastCol; c++)
                    {
                        width += columnWidths[c];
                    }
                    
                    height = 0;
                    for (int r = mergeFirstRow; r <= mergeLastRow && r <= lastRow; r++)
                    {
                        height += rowHeights[r];
                    }
                    
                    // Mark all cells in the merged range as drawn
                    for (int r = mergeFirstRow; r <= mergeLastRow && r <= lastRow; r++)
                    {
                        for (int c = mergeFirstCol; c <= mergeLastCol && c <= lastCol; c++)
                        {
                            drawnCells.Add($"{r},{c}");
                        }
                    }
                }
                
                DrawCell(canvas, cell, x, y, width, height);
            }
        }
        
        // Render to PNG
        using var image = surface.Snapshot();
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        return data.ToArray();
    }
    
    /// <summary>
    /// Render all worksheets in a workbook to individual PNG images
    /// </summary>
    public Dictionary<string, byte[]> RenderWorkbookToPngs(XLWorkbook workbook)
    {
        var images = new Dictionary<string, byte[]>();
        
        foreach (var worksheet in workbook.Worksheets)
        {
            var imageData = RenderWorksheetToPng(worksheet);
            images[worksheet.Name] = imageData;
        }
        
        return images;
    }
    
    /// <summary>
    /// Render all worksheets and combine into a single image
    /// </summary>
    public byte[] RenderWorkbookToSinglePng(XLWorkbook workbook)
    {
        var worksheetImages = new List<(string name, byte[] data)>();
        
        foreach (var worksheet in workbook.Worksheets)
        {
            var imageData = RenderWorksheetToPng(worksheet);
            worksheetImages.Add((worksheet.Name, imageData));
        }
        
        if (worksheetImages.Count == 0)
        {
            return CreateEmptyWorksheetImage();
        }
        
        // Calculate total height (stack worksheets vertically)
        int totalHeight = 0;
        int maxWidth = 0;
        var decodedImages = new List<SKBitmap>();
        
        foreach (var (name, data) in worksheetImages)
        {
            using var ms = new MemoryStream(data);
            var bitmap = SKBitmap.Decode(ms);
            decodedImages.Add(bitmap);
            totalHeight += bitmap.Height + 10; // Add spacing between worksheets
            maxWidth = Math.Max(maxWidth, bitmap.Width);
        }
        
        // Create combined image
        var imageInfo = new SKImageInfo(maxWidth, totalHeight);
        using var surface = SKSurface.Create(imageInfo);
        var canvas = surface.Canvas;
        canvas.Clear(SKColors.LightGray);
        
        int currentY = 0;
        foreach (var bitmap in decodedImages)
        {
            canvas.DrawBitmap(bitmap, 0, currentY);
            currentY += bitmap.Height + 10;
            bitmap.Dispose();
        }
        
        using var image = surface.Snapshot();
        using var finalData = image.Encode(SKEncodedImageFormat.Png, 100);
        return finalData.ToArray();
    }
    
    private void DrawWorksheetHeader(SKCanvas canvas, string worksheetName, int width)
    {
        using var headerPaint = new SKPaint
        {
            Color = SKColors.DarkBlue,
            IsAntialias = true,
            Style = SKPaintStyle.Fill
        };
        
        canvas.DrawRect(0, 0, width, WorksheetNameHeight, headerPaint);
        
        using var textPaint = new SKPaint
        {
            Color = SKColors.White,
            IsAntialias = true,
            TextSize = 16,
            Typeface = SKTypeface.FromFamilyName("Arial", SKFontStyle.Bold)
        };
        
        canvas.DrawText(worksheetName, 10, 20, textPaint);
    }
    
    private void DrawColumnHeaders(SKCanvas canvas, int firstCol, int lastCol, Dictionary<int, double> columnPositions, Dictionary<int, double> columnWidths)
    {
        using var bgPaint = new SKPaint
        {
            Color = new SKColor(242, 242, 242), // Light grey background like Excel
            Style = SKPaintStyle.Fill
        };
        
        using var borderPaint = new SKPaint
        {
            Color = new SKColor(217, 217, 217),
            Style = SKPaintStyle.Stroke,
            StrokeWidth = 1
        };
        
        using var textPaint = new SKPaint
        {
            Color = SKColors.Black,
            IsAntialias = true,
            TextSize = 10,
            Typeface = SKTypeface.FromFamilyName("Segoe UI")
        };
        
        var headerY = WorksheetNameHeight;
        
        for (int col = firstCol; col <= lastCol; col++)
        {
            var x = columnPositions[col];
            var width = columnWidths[col];
            
            // Draw background
            canvas.DrawRect((float)x, (float)headerY, (float)width, ColumnHeaderHeight, bgPaint);
            
            // Draw border
            canvas.DrawRect((float)x, (float)headerY, (float)width, ColumnHeaderHeight, borderPaint);
            
            // Draw column letter
            var columnLetter = GetColumnLetter(col);
            var textWidth = textPaint.MeasureText(columnLetter);
            var textX = (float)(x + (width - textWidth) / 2);
            var textY = (float)(headerY + ColumnHeaderHeight / 2 + 4);
            canvas.DrawText(columnLetter, textX, textY, textPaint);
        }
    }
    
    private void DrawRowHeaders(SKCanvas canvas, int firstRow, int lastRow, Dictionary<int, double> rowPositions, Dictionary<int, double> rowHeights)
    {
        using var bgPaint = new SKPaint
        {
            Color = new SKColor(242, 242, 242), // Light grey background like Excel
            Style = SKPaintStyle.Fill
        };
        
        using var borderPaint = new SKPaint
        {
            Color = new SKColor(217, 217, 217),
            Style = SKPaintStyle.Stroke,
            StrokeWidth = 1
        };
        
        using var textPaint = new SKPaint
        {
            Color = SKColors.Black,
            IsAntialias = true,
            TextSize = 10,
            Typeface = SKTypeface.FromFamilyName("Segoe UI")
        };
        
        for (int row = firstRow; row <= lastRow; row++)
        {
            var y = rowPositions[row];
            var height = rowHeights[row];
            
            // Draw background
            canvas.DrawRect(0, (float)y, RowHeaderWidth, (float)height, bgPaint);
            
            // Draw border
            canvas.DrawRect(0, (float)y, RowHeaderWidth, (float)height, borderPaint);
            
            // Draw row number
            var rowNumber = row.ToString();
            var textWidth = textPaint.MeasureText(rowNumber);
            var textX = (RowHeaderWidth - textWidth) / 2;
            var textY = (float)(y + height / 2 + 4);
            canvas.DrawText(rowNumber, textX, textY, textPaint);
        }
    }
    
    private string GetColumnLetter(int columnNumber)
    {
        string columnLetter = "";
        while (columnNumber > 0)
        {
            int modulo = (columnNumber - 1) % 26;
            columnLetter = Convert.ToChar('A' + modulo) + columnLetter;
            columnNumber = (columnNumber - modulo) / 26;
        }
        return columnLetter;
    }
    
    private void DrawCell(SKCanvas canvas, IXLCell cell, double x, double y, double width, double height)
    {
        // Determine cell background color
        var bgColor = SKColors.White;
        if (cell.Style.Fill.BackgroundColor.ColorType == XLColorType.Color)
        {
            var xlColor = cell.Style.Fill.BackgroundColor.Color;
            bgColor = new SKColor(xlColor.R, xlColor.G, xlColor.B);
        }
        
        // Draw cell background
        using var bgPaint = new SKPaint
        {
            Color = bgColor,
            Style = SKPaintStyle.Fill
        };
        canvas.DrawRect((float)x, (float)y, (float)width, (float)height, bgPaint);
        
        // Draw cell borders (all four sides individually for proper styling)
        DrawCellBorders(canvas, cell, x, y, width, height);
        
        // Draw cell text
        var cellText = GetCellDisplayText(cell);
        if (!string.IsNullOrEmpty(cellText))
        {
            var textColor = SKColors.Black;
            if (cell.Style.Font.FontColor.ColorType == XLColorType.Color)
            {
                var xlColor = cell.Style.Font.FontColor.Color;
                textColor = new SKColor(xlColor.R, xlColor.G, xlColor.B);
            }
            
            // Get font size from cell style
            var fontSize = (float)cell.Style.Font.FontSize;
            if (fontSize < 6) fontSize = 10; // Default if not set
            
            // Determine font style
            var fontStyle = SKFontStyle.Normal;
            if (cell.Style.Font.Bold && cell.Style.Font.Italic)
                fontStyle = SKFontStyle.BoldItalic;
            else if (cell.Style.Font.Bold)
                fontStyle = SKFontStyle.Bold;
            else if (cell.Style.Font.Italic)
                fontStyle = SKFontStyle.Italic;
            
            // Use a font that supports more Unicode characters
            var typeface = SKTypeface.FromFamilyName("Segoe UI", fontStyle) 
                ?? SKTypeface.FromFamilyName("Arial Unicode MS", fontStyle)
                ?? SKTypeface.FromFamilyName("Arial", fontStyle);
            
            using var textPaint = new SKPaint
            {
                Color = textColor,
                IsAntialias = true,
                TextSize = fontSize,
                Typeface = typeface,
                TextEncoding = SKTextEncoding.Utf8
            };
            
            // Calculate text position (centered vertically in cell)
            var textHeight = textPaint.FontMetrics.Descent - textPaint.FontMetrics.Ascent;
            var textY = (float)(y + (height + textHeight) / 2 - textPaint.FontMetrics.Descent);
            
            // Truncate text if too long
            var truncatedText = TruncateText(cellText, width - 10, textPaint);
            canvas.DrawText(truncatedText, (float)(x + 5), textY, textPaint);
        }
    }
    
    private void DrawCellBorders(SKCanvas canvas, IXLCell cell, double x, double y, double width, double height)
    {
        var hasBorders = cell.Style.Border.TopBorder != XLBorderStyleValues.None ||
                        cell.Style.Border.BottomBorder != XLBorderStyleValues.None ||
                        cell.Style.Border.LeftBorder != XLBorderStyleValues.None ||
                        cell.Style.Border.RightBorder != XLBorderStyleValues.None;
        
        if (!hasBorders)
        {
            // Draw light grey gridlines if no borders are set
            using var gridPaint = new SKPaint
            {
                Color = new SKColor(217, 217, 217), // Light grey like Excel gridlines
                Style = SKPaintStyle.Stroke,
                StrokeWidth = 1,
                IsAntialias = false
            };
            canvas.DrawRect((float)x, (float)y, (float)width, (float)height, gridPaint);
        }
        else
        {
            // Draw each border side individually to support different styles and colors
            DrawBorder(canvas, cell.Style.Border.TopBorder, cell.Style.Border.TopBorderColor, x, y, x + width, y); // Top
            DrawBorder(canvas, cell.Style.Border.BottomBorder, cell.Style.Border.BottomBorderColor, x, y + height, x + width, y + height); // Bottom
            DrawBorder(canvas, cell.Style.Border.LeftBorder, cell.Style.Border.LeftBorderColor, x, y, x, y + height); // Left
            DrawBorder(canvas, cell.Style.Border.RightBorder, cell.Style.Border.RightBorderColor, x + width, y, x + width, y + height); // Right
        }
    }
    
    private void DrawBorder(SKCanvas canvas, XLBorderStyleValues borderStyle, XLColor borderColor, double x1, double y1, double x2, double y2)
    {
        if (borderStyle == XLBorderStyleValues.None)
            return;
        
        // Get border color
        var skBorderColor = SKColors.Black;
        if (borderColor.ColorType == XLColorType.Color)
        {
            var xlColor = borderColor.Color;
            skBorderColor = new SKColor(xlColor.R, xlColor.G, xlColor.B);
        }
        
        // Determine line width based on border style
        float strokeWidth = borderStyle switch
        {
            XLBorderStyleValues.Thin => 1f,
            XLBorderStyleValues.Medium => 2f,
            XLBorderStyleValues.Thick => 3f,
            XLBorderStyleValues.Hair => 0.5f,
            _ => 1f
        };
        
        using var paint = new SKPaint
        {
            Color = skBorderColor,
            Style = SKPaintStyle.Stroke,
            StrokeWidth = strokeWidth,
            IsAntialias = true
        };
        
        // Handle different border styles
        switch (borderStyle)
        {
            case XLBorderStyleValues.Dotted:
                paint.PathEffect = SKPathEffect.CreateDash(new float[] { 2, 2 }, 0);
                canvas.DrawLine((float)x1, (float)y1, (float)x2, (float)y2, paint);
                break;
                
            case XLBorderStyleValues.Dashed:
            case XLBorderStyleValues.MediumDashed:
                paint.PathEffect = SKPathEffect.CreateDash(new float[] { 5, 3 }, 0);
                canvas.DrawLine((float)x1, (float)y1, (float)x2, (float)y2, paint);
                break;
                
            case XLBorderStyleValues.DashDot:
            case XLBorderStyleValues.MediumDashDot:
                paint.PathEffect = SKPathEffect.CreateDash(new float[] { 5, 3, 2, 3 }, 0);
                canvas.DrawLine((float)x1, (float)y1, (float)x2, (float)y2, paint);
                break;
                
            case XLBorderStyleValues.DashDotDot:
            case XLBorderStyleValues.MediumDashDotDot:
                paint.PathEffect = SKPathEffect.CreateDash(new float[] { 5, 3, 2, 3, 2, 3 }, 0);
                canvas.DrawLine((float)x1, (float)y1, (float)x2, (float)y2, paint);
                break;
                
            case XLBorderStyleValues.Double:
                // Draw two parallel lines for double border with thinner strokes
                paint.StrokeWidth = 1f; // Use thinner lines for double borders
                var offset = 3f; // Distance between the two lines
                bool isHorizontal = Math.Abs(y2 - y1) < 0.1;
                
                if (isHorizontal)
                {
                    // Horizontal line - offset vertically
                    canvas.DrawLine((float)x1, (float)y1 - offset, (float)x2, (float)y2 - offset, paint);
                    canvas.DrawLine((float)x1, (float)y1 - 1f, (float)x2, (float)y2 - 1f, paint);
                }
                else
                {
                    // Vertical line - offset horizontally
                    canvas.DrawLine((float)x1 - offset, (float)y1, (float)x2 - offset, (float)y2, paint);
                    canvas.DrawLine((float)x1 + offset, (float)y1, (float)x2 + offset, (float)y2, paint);
                }
                break;
                
            default:
                // Standard solid line for other styles
                canvas.DrawLine((float)x1, (float)y1, (float)x2, (float)y2, paint);
                break;
        }
    }
    
    private string GetCellDisplayText(IXLCell cell)
    {
        // Show formula errors if present
        if (cell.HasFormula && cell.Value.IsError)
        {
            return $"#ERROR";
        }
        
        // Display calculated value
        if (cell.Value.IsNumber)
        {
            // Try to respect number format
            return cell.GetFormattedString();
        }
        
        return cell.Value.ToString() ?? "";
    }
    
    private string TruncateText(string text, double maxWidth, SKPaint paint)
    {
        if (string.IsNullOrEmpty(text))
            return text;
        
        var textWidth = paint.MeasureText(text);
        if (textWidth <= maxWidth)
            return text;
        
        // Truncate and add ellipsis
        var ellipsis = "...";
        var ellipsisWidth = paint.MeasureText(ellipsis);
        var availableWidth = maxWidth - ellipsisWidth;
        
        for (int i = text.Length - 1; i > 0; i--)
        {
            var substring = text.Substring(0, i);
            if (paint.MeasureText(substring) <= availableWidth)
            {
                return substring + ellipsis;
            }
        }
        
        return ellipsis;
    }
    
    private byte[] CreateEmptyWorksheetImage()
    {
        var imageInfo = new SKImageInfo(400, 100);
        using var surface = SKSurface.Create(imageInfo);
        var canvas = surface.Canvas;
        
        canvas.Clear(SKColors.White);
        
        using var textPaint = new SKPaint
        {
            Color = SKColors.Gray,
            IsAntialias = true,
            TextSize = 14,
            Typeface = SKTypeface.FromFamilyName("Arial")
        };
        
        canvas.DrawText("(Empty Worksheet)", 100, 50, textPaint);
        
        using var image = surface.Snapshot();
        using var data = image.Encode(SKEncodedImageFormat.Png, 100);
        return data.ToArray();
    }
}
