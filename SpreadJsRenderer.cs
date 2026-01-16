using GrapeCity.Documents.Excel;
using SkiaSharp;

namespace ExcelAgent.Services;

/// <summary>
/// Renders SpreadJS (GrapeCity Documents.Excel) worksheets to images using SkiaSharp
/// </summary>
public class SpreadJSRenderer
{
    private const int WorksheetNameHeight = 30;
    private const int ColumnHeaderHeight = 20;
    private const int RowHeaderWidth = 40;
    private const int MaxColumnsToRender = 100;
    private const int MaxRowsToRender = 100;
    private const double PixelsPerExcelWidth = 7.5; // Excel column width units to pixels
    private const double DefaultColumnWidth = 64.0; // Default column width in pixels
    private const double DefaultRowHeight = 20.0; // Default row height in pixels
    
    /// <summary>
    /// Render a worksheet to a PNG image
    /// </summary>
    public byte[] RenderWorksheetToPng(IWorksheet worksheet, int maxColumns = MaxColumnsToRender, int maxRows = MaxRowsToRender)
    {
        var usedRange = worksheet.UsedRange;
        if (usedRange == null || usedRange.RowCount == 0 || usedRange.ColumnCount == 0)
        {
            // Empty worksheet - create a small placeholder image
            return CreateEmptyWorksheetImage();
        }
        
        var firstRow = 0; // SpreadJS uses 0-based indexing
        var firstCol = 0;
        var lastRow = Math.Min(usedRange.Row + usedRange.RowCount - 1, firstRow + maxRows - 1);
        var lastCol = Math.Min(usedRange.Column + usedRange.ColumnCount - 1, firstCol + maxColumns - 1);
        
        // Calculate actual column widths and row heights
        var columnWidths = new Dictionary<int, double>();
        var rowHeights = new Dictionary<int, double>();
        var columnPositions = new Dictionary<int, double>();
        var rowPositions = new Dictionary<int, double>();
        
        double totalWidth = RowHeaderWidth; // Start after row header
        for (int col = firstCol; col <= lastCol; col++)
        {
            var width = worksheet.Columns[col].ColumnWidth * PixelsPerExcelWidth;
            if (width < 1) width = DefaultColumnWidth; // Use default if width is 0 or very small
            columnWidths[col] = width;
            columnPositions[col] = totalWidth;
            totalWidth += width;
        }
        
        double totalHeight = WorksheetNameHeight + ColumnHeaderHeight; // Start after worksheet name and column headers
        for (int row = firstRow; row <= lastRow; row++)
        {
            var height = worksheet.Rows[row].RowHeight;
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
                var cell = worksheet.Range[row, col];
                var cellKey = $"{row},{col}";
                
                // Skip if already drawn as part of a merged cell
                if (drawnCells.Contains(cellKey))
                    continue;
                
                var x = columnPositions[col];
                var y = rowPositions[row];
                var width = columnWidths[col];
                var height = rowHeights[row];
                
                // Check if this is a merged cell
                if (cell.MergeCells)
                {
                    var mergeArea = cell.MergeArea;
                    if (mergeArea != null)
                    {
                        var mergeFirstRow = mergeArea.Row;
                        var mergeFirstCol = mergeArea.Column;
                        var mergeLastRow = mergeArea.Row + mergeArea.RowCount - 1;
                        var mergeLastCol = mergeArea.Column + mergeArea.ColumnCount - 1;
                        
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
    public Dictionary<string, byte[]> RenderWorkbookToPngs(Workbook workbook)
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
    public byte[] RenderWorkbookToSinglePng(Workbook workbook)
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
            
            // Draw column letter (SpreadJS uses 0-based, display as 1-based like Excel)
            var columnLetter = GetColumnLetter(col + 1);
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
            
            // Draw row number (SpreadJS uses 0-based, display as 1-based like Excel)
            var rowNumber = (row + 1).ToString();
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
    
    private void DrawCell(SKCanvas canvas, IRange cell, double x, double y, double width, double height)
    {
        // Determine cell background color
        var bgColor = SKColors.White;
        var color = cell.Interior.Color;
        if (color != System.Drawing.Color.Empty && color.A > 0)
        {
            bgColor = ColorToSKColor(color);
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
            var fontColor = cell.Font.Color;
            if (fontColor != System.Drawing.Color.Empty && fontColor.A > 0)
            {
                textColor = ColorToSKColor(fontColor);
            }
            
            // Get font size from cell style
            var fontSize = (float)cell.Font.Size;
            if (fontSize < 6) fontSize = 10; // Default if not set
            
            // Determine font style
            var fontStyle = SKFontStyle.Normal;
            if (cell.Font.Bold && cell.Font.Italic)
                fontStyle = SKFontStyle.BoldItalic;
            else if (cell.Font.Bold)
                fontStyle = SKFontStyle.Bold;
            else if (cell.Font.Italic)
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
    
    private void DrawCellBorders(SKCanvas canvas, IRange cell, double x, double y, double width, double height)
    {
        var borders = cell.Borders;
        var hasBorders = borders[BordersIndex.EdgeTop].LineStyle != BorderLineStyle.None ||
                        borders[BordersIndex.EdgeBottom].LineStyle != BorderLineStyle.None ||
                        borders[BordersIndex.EdgeLeft].LineStyle != BorderLineStyle.None ||
                        borders[BordersIndex.EdgeRight].LineStyle != BorderLineStyle.None;
        
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
            DrawBorder(canvas, borders[BordersIndex.EdgeTop], x, y, x + width, y); // Top
            DrawBorder(canvas, borders[BordersIndex.EdgeBottom], x, y + height, x + width, y + height); // Bottom
            DrawBorder(canvas, borders[BordersIndex.EdgeLeft], x, y, x, y + height); // Left
            DrawBorder(canvas, borders[BordersIndex.EdgeRight], x + width, y, x + width, y + height); // Right
        }
    }
    
    private void DrawBorder(SKCanvas canvas, IBorder border, double x1, double y1, double x2, double y2)
    {
        if (border.LineStyle == BorderLineStyle.None)
            return;
        
        // Get border color
        var skBorderColor = SKColors.Black;
        var borderColor = border.Color;
        if (borderColor != System.Drawing.Color.Empty && borderColor.A > 0)
        {
            skBorderColor = ColorToSKColor(borderColor);
        }
        
        // Determine line width based on border style
        float strokeWidth = border.LineStyle switch
        {
            BorderLineStyle.Hair => 0.5f,
            BorderLineStyle.Thin => 1f,
            BorderLineStyle.Medium => 2f,
            BorderLineStyle.Thick => 3f,
            BorderLineStyle.MediumDashDot => 2f,
            BorderLineStyle.MediumDashDotDot => 2f,
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
        switch (border.LineStyle)
        {
            case BorderLineStyle.Dotted:
                paint.PathEffect = SKPathEffect.CreateDash(new float[] { 2, 2 }, 0);
                canvas.DrawLine((float)x1, (float)y1, (float)x2, (float)y2, paint);
                break;
                
            case BorderLineStyle.Dashed:
                paint.PathEffect = SKPathEffect.CreateDash(new float[] { 5, 3 }, 0);
                canvas.DrawLine((float)x1, (float)y1, (float)x2, (float)y2, paint);
                break;
                
            case BorderLineStyle.DashDot:
            case BorderLineStyle.MediumDashDot:
                paint.PathEffect = SKPathEffect.CreateDash(new float[] { 5, 3, 2, 3 }, 0);
                canvas.DrawLine((float)x1, (float)y1, (float)x2, (float)y2, paint);
                break;
                
            case BorderLineStyle.DashDotDot:
            case BorderLineStyle.MediumDashDotDot:
                paint.PathEffect = SKPathEffect.CreateDash(new float[] { 5, 3, 2, 3, 2, 3 }, 0);
                canvas.DrawLine((float)x1, (float)y1, (float)x2, (float)y2, paint);
                break;
                
            case BorderLineStyle.Double:
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
    
    private string GetCellDisplayText(IRange cell)
    {
        // Show formula errors if present
        if (cell.HasFormula && cell.Value != null && cell.Value.ToString().StartsWith("#"))
        {
            return cell.Value.ToString();
        }
        
        // Display calculated value with number formatting
        if (cell.Value != null)
        {
            return cell.Text; // Use formatted text representation
        }
        
        return "";
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
    
    private SKColor ColorToSKColor(System.Drawing.Color color)
    {
        // System.Drawing.Color to SkiaSharp color conversion
        return new SKColor(color.R, color.G, color.B, color.A);
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
