using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Wordprocessing;

namespace DecoSOP.Services;

/// <summary>
/// Generates HTML previews from Office document files (DOCX, XLSX, DOC).
/// </summary>
public static class DocumentPreviewService
{
    private static readonly HashSet<string> PreviewableExtensions =
        [".pdf", ".png", ".jpg", ".jpeg", ".gif", ".bmp", ".webp", ".docx", ".xlsx", ".doc"];

    public static bool CanPreview(string fileName)
    {
        var ext = Path.GetExtension(fileName).ToLowerInvariant();
        return PreviewableExtensions.Contains(ext);
    }

    public static bool NeedsHtmlConversion(string fileName)
    {
        var ext = Path.GetExtension(fileName).ToLowerInvariant();
        return ext is ".docx" or ".xlsx" or ".doc";
    }

    /// <summary>
    /// Convert a file to an HTML preview. Returns a full HTML document string.
    /// </summary>
    public static string GenerateHtmlPreview(string filePath, string title)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        var body = ext switch
        {
            ".docx" => ConvertDocxToHtml(filePath),
            ".xlsx" => ConvertXlsxToHtml(filePath),
            ".doc" => ConvertDocToHtml(filePath),
            _ => "<p>Preview not available for this file type.</p>"
        };

        var safeTitle = HttpUtility.HtmlEncode(title);
        return $$"""
            <!DOCTYPE html>
            <html>
            <head>
            <meta charset="utf-8" />
            <title>{{safeTitle}}</title>
            <style>
                body { font-family: 'Segoe UI', Arial, sans-serif; margin: 2rem; color: #333; line-height: 1.6; }
                h1 { color: #1b6ec2; font-size: 1.5rem; border-bottom: 2px solid #e9ecef; padding-bottom: 0.5rem; }
                table { border-collapse: collapse; width: 100%; margin: 1rem 0; }
                th, td { border: 1px solid #dee2e6; padding: 0.5rem 0.75rem; text-align: left; vertical-align: top; }
                th { background: #f0f4f8; font-weight: 600; }
                p { margin: 0.3rem 0; }
                .sheet-title { color: #555; font-size: 1.1rem; margin-top: 1.5rem; }
                @media print {
                    body { margin: 0.5rem; }
                    table { font-size: 0.85rem; }
                }
            </style>
            </head>
            <body>
            <h1>{{safeTitle}}</h1>
            {{body}}
            </body>
            </html>
            """;
    }

    // ============================================================
    //  DOCX conversion (OpenXml)
    // ============================================================

    private static string ConvertDocxToHtml(string filePath)
    {
        try
        {
            using var doc = WordprocessingDocument.Open(filePath, false);
            var body = doc.MainDocumentPart?.Document?.Body;
            if (body is null) return "<p>Unable to read document content.</p>";

            var sb = new StringBuilder();
            foreach (var element in body.ChildElements)
            {
                if (element is Paragraph para)
                {
                    var text = GetParagraphText(para);
                    if (string.IsNullOrWhiteSpace(text)) continue;

                    var isBold = IsParagraphBold(para);
                    var tag = isBold && text.Length < 120 ? "h2" : "p";
                    sb.AppendLine($"<{tag}>{HttpUtility.HtmlEncode(text)}</{tag}>");
                }
                else if (element is DocumentFormat.OpenXml.Wordprocessing.Table table)
                {
                    sb.AppendLine(ConvertWordTableToHtml(table));
                }
            }

            return sb.Length > 0 ? sb.ToString() : "<p>Document appears to be empty.</p>";
        }
        catch (Exception ex)
        {
            return $"<p>Unable to preview this document: {HttpUtility.HtmlEncode(ex.Message)}</p>";
        }
    }

    private static string GetParagraphText(Paragraph para)
    {
        var sb = new StringBuilder();
        foreach (var run in para.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>())
        {
            foreach (var text in run.Elements<DocumentFormat.OpenXml.Wordprocessing.Text>())
                sb.Append(text.Text);
        }
        return sb.ToString();
    }

    private static bool IsParagraphBold(Paragraph para)
    {
        var pProps = para.ParagraphProperties;
        if (pProps?.ParagraphMarkRunProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Bold>() is not null)
            return true;

        var runs = para.Elements<DocumentFormat.OpenXml.Wordprocessing.Run>().ToList();
        if (runs.Count == 0) return false;
        return runs.All(r => r.RunProperties?.GetFirstChild<DocumentFormat.OpenXml.Wordprocessing.Bold>() is not null);
    }

    private static string ConvertWordTableToHtml(DocumentFormat.OpenXml.Wordprocessing.Table table)
    {
        var sb = new StringBuilder();
        sb.AppendLine("<table>");

        var rows = table.Elements<TableRow>().ToList();
        for (int i = 0; i < rows.Count; i++)
        {
            sb.AppendLine("<tr>");
            var cellTag = i == 0 ? "th" : "td";
            foreach (var cell in rows[i].Elements<TableCell>())
            {
                var text = cell.InnerText;
                sb.AppendLine($"<{cellTag}>{HttpUtility.HtmlEncode(text)}</{cellTag}>");
            }
            sb.AppendLine("</tr>");
        }

        sb.AppendLine("</table>");
        return sb.ToString();
    }

    // ============================================================
    //  XLSX conversion (OpenXml) — with cell formatting
    // ============================================================

    // Standard Excel indexed color palette
    private static readonly string[] IndexedColors =
    [
        "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
        "#000000", "#FFFFFF", "#FF0000", "#00FF00", "#0000FF", "#FFFF00", "#FF00FF", "#00FFFF",
        "#800000", "#008000", "#000080", "#808000", "#800080", "#008080", "#C0C0C0", "#808080",
        "#9999FF", "#993366", "#FFFFCC", "#CCFFFF", "#660066", "#FF8080", "#0066CC", "#CCCCFF",
        "#000080", "#FF00FF", "#FFFF00", "#00FFFF", "#800080", "#800000", "#008080", "#0000FF",
        "#00CCFF", "#CCFFFF", "#CCFFCC", "#FFFF99", "#99CCFF", "#FF99CC", "#CC99FF", "#FFCC99",
        "#3366FF", "#33CCCC", "#99CC00", "#FFCC00", "#FF9900", "#FF6600", "#666699", "#969696",
        "#003366", "#339966", "#003300", "#333300", "#993300", "#993366", "#333399", "#333333",
    ];

    private static string ConvertXlsxToHtml(string filePath)
    {
        try
        {
            using var doc = SpreadsheetDocument.Open(filePath, false);
            var workbookPart = doc.WorkbookPart;
            if (workbookPart is null) return "<p>Unable to read spreadsheet.</p>";

            var sharedStrings = workbookPart.SharedStringTablePart?.SharedStringTable?
                .Elements<SharedStringItem>()
                .Select(s => s.InnerText)
                .ToList() ?? [];

            var stylesheet = workbookPart.WorkbookStylesPart?.Stylesheet;
            var themeColors = GetThemeColors(workbookPart);

            var sheets = workbookPart.Workbook?.Sheets?.Elements<Sheet>().ToList() ?? [];
            if (sheets.Count == 0) return "<p>Spreadsheet has no sheets.</p>";

            var sb = new StringBuilder();

            foreach (var sheet in sheets)
            {
                var worksheetPart = (WorksheetPart?)workbookPart.GetPartById(sheet.Id!);
                if (worksheetPart is null) continue;

                var sheetData = worksheetPart.Worksheet?.GetFirstChild<SheetData>();
                if (sheetData is null) continue;

                var rows = sheetData.Elements<Row>().ToList();
                if (rows.Count == 0) continue;

                if (sheets.Count > 1)
                    sb.AppendLine($"<h2 class=\"sheet-title\">{HttpUtility.HtmlEncode(sheet.Name?.Value ?? "Sheet")}</h2>");

                sb.AppendLine("<table>");
                for (int i = 0; i < rows.Count; i++)
                {
                    sb.AppendLine("<tr>");
                    foreach (var cell in rows[i].Elements<Cell>())
                    {
                        var value = GetCellValue(cell, sharedStrings);
                        var style = GetCellStyleCss(cell, stylesheet, themeColors);
                        var tag = i == 0 ? "th" : "td";

                        if (!string.IsNullOrEmpty(style))
                            sb.AppendLine($"<{tag} style=\"{style}\">{HttpUtility.HtmlEncode(value)}</{tag}>");
                        else
                            sb.AppendLine($"<{tag}>{HttpUtility.HtmlEncode(value)}</{tag}>");
                    }
                    sb.AppendLine("</tr>");
                }
                sb.AppendLine("</table>");
            }

            return sb.Length > 0 ? sb.ToString() : "<p>Spreadsheet appears to be empty.</p>";
        }
        catch (Exception ex)
        {
            return $"<p>Unable to preview this spreadsheet: {HttpUtility.HtmlEncode(ex.Message)}</p>";
        }
    }

    private static string GetCellValue(Cell cell, List<string> sharedStrings)
    {
        var value = cell.CellValue?.Text ?? "";
        if (cell.DataType?.Value == CellValues.SharedString
            && int.TryParse(value, out var index)
            && index >= 0 && index < sharedStrings.Count)
        {
            return sharedStrings[index];
        }
        return value;
    }

    private static string GetCellStyleCss(Cell cell, Stylesheet? stylesheet, List<string> themeColors)
    {
        if (stylesheet is null || cell.StyleIndex is null) return "";

        var styleIndex = (int)cell.StyleIndex.Value;
        var cellFormats = stylesheet.CellFormats;
        if (cellFormats is null) return "";

        var cellFormat = cellFormats.Elements<CellFormat>().ElementAtOrDefault(styleIndex);
        if (cellFormat is null) return "";

        var css = new List<string>();

        // Background color from fill
        if (cellFormat.FillId is not null)
        {
            var fillIndex = (int)cellFormat.FillId.Value;
            var fill = stylesheet.Fills?.Elements<Fill>().ElementAtOrDefault(fillIndex);
            var patternFill = fill?.PatternFill;
            if (patternFill?.PatternType?.Value == PatternValues.Solid)
            {
                var bgColor = ResolveColor(patternFill.ForegroundColor, themeColors);
                if (bgColor is not null && bgColor != "#FFFFFF" && bgColor != "#000000")
                    css.Add($"background-color:{bgColor}");
                // If bg is dark, make text white automatically
                if (bgColor is not null && IsDarkColor(bgColor))
                    css.Add("color:#FFFFFF");
            }
        }

        // Font styles
        if (cellFormat.FontId is not null)
        {
            var fontIndex = (int)cellFormat.FontId.Value;
            var font = stylesheet.Fonts?.Elements<DocumentFormat.OpenXml.Spreadsheet.Font>()
                .ElementAtOrDefault(fontIndex);
            if (font is not null)
            {
                var fontColor = ResolveColor(font.Color, themeColors);
                if (fontColor is not null && fontColor != "#000000")
                    css.Add($"color:{fontColor}");

                if (font.Bold is not null)
                    css.Add("font-weight:bold");

                if (font.Italic is not null)
                    css.Add("font-style:italic");

                if (font.FontSize?.Val is not null)
                {
                    var size = font.FontSize.Val.Value;
                    if (size > 11) // only set if larger than default
                        css.Add($"font-size:{size}pt");
                }
            }
        }

        // Alignment
        if (cellFormat.Alignment is not null)
        {
            var align = cellFormat.Alignment.Horizontal?.Value;
            if (align == DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Center)
                css.Add("text-align:center");
            else if (align == DocumentFormat.OpenXml.Spreadsheet.HorizontalAlignmentValues.Right)
                css.Add("text-align:right");
        }

        return string.Join(";", css);
    }

    private static string? ResolveColor(ColorType? color, List<string> themeColors)
    {
        if (color is null) return null;

        // Direct RGB (ARGB format — skip alpha channel)
        if (color.Rgb is not null)
        {
            var rgb = color.Rgb.Value ?? "";
            if (rgb.Length == 8) rgb = rgb[2..]; // strip alpha
            if (rgb.Length == 6) return $"#{rgb}";
        }

        // Theme color with optional tint
        if (color.Theme is not null)
        {
            var themeIndex = (int)color.Theme.Value;
            if (themeIndex < themeColors.Count)
            {
                var baseColor = themeColors[themeIndex];
                if (color.Tint?.Value is double tint and not 0)
                    return ApplyTint(baseColor, tint);
                return baseColor;
            }
        }

        // Indexed color
        if (color.Indexed is not null)
        {
            var idx = (int)color.Indexed.Value;
            if (idx >= 0 && idx < IndexedColors.Length)
                return IndexedColors[idx];
        }

        return null;
    }

    private static List<string> GetThemeColors(WorkbookPart workbookPart)
    {
        var colors = new List<string>();
        try
        {
            var themePart = workbookPart.ThemePart;
            if (themePart?.Theme?.ThemeElements?.ColorScheme is null)
                return colors;

            var cs = themePart.Theme.ThemeElements.ColorScheme;

            // Excel theme color order: dk1, lt1, dk2, lt2, accent1-6, hyperlink, followed
            // Each Color2Type has RgbColorModelHex or SystemColor
            string? GetHex(DocumentFormat.OpenXml.Drawing.Color2Type? c)
            {
                if (c is null) return null;
                return c.RgbColorModelHex?.Val?.Value ?? c.SystemColor?.LastColor?.Value;
            }

            var themeEntries = new[]
            {
                GetHex(cs.Dark1Color), GetHex(cs.Light1Color),
                GetHex(cs.Dark2Color), GetHex(cs.Light2Color),
                GetHex(cs.Accent1Color), GetHex(cs.Accent2Color),
                GetHex(cs.Accent3Color), GetHex(cs.Accent4Color),
                GetHex(cs.Accent5Color), GetHex(cs.Accent6Color),
                GetHex(cs.Hyperlink), GetHex(cs.FollowedHyperlinkColor)
            };

            foreach (var hex in themeEntries)
            {
                colors.Add(hex is not null ? $"#{hex}" : "#000000");
            }
        }
        catch { }
        return colors;
    }

    private static string ApplyTint(string hexColor, double tint)
    {
        // Parse hex color
        var hex = hexColor.TrimStart('#');
        if (hex.Length != 6) return hexColor;

        var r = Convert.ToInt32(hex[..2], 16);
        var g = Convert.ToInt32(hex[2..4], 16);
        var b = Convert.ToInt32(hex[4..6], 16);

        if (tint > 0)
        {
            // Lighten: mix towards white
            r = (int)(r + (255 - r) * tint);
            g = (int)(g + (255 - g) * tint);
            b = (int)(b + (255 - b) * tint);
        }
        else
        {
            // Darken: mix towards black
            r = (int)(r * (1 + tint));
            g = (int)(g * (1 + tint));
            b = (int)(b * (1 + tint));
        }

        r = Math.Clamp(r, 0, 255);
        g = Math.Clamp(g, 0, 255);
        b = Math.Clamp(b, 0, 255);

        return $"#{r:X2}{g:X2}{b:X2}";
    }

    private static bool IsDarkColor(string hexColor)
    {
        var hex = hexColor.TrimStart('#');
        if (hex.Length != 6) return false;
        var r = Convert.ToInt32(hex[..2], 16);
        var g = Convert.ToInt32(hex[2..4], 16);
        var b = Convert.ToInt32(hex[4..6], 16);
        // Perceived luminance
        return (0.299 * r + 0.587 * g + 0.114 * b) < 128;
    }

    // ============================================================
    //  DOC conversion (binary Word 97-2003 / XML Word 2003)
    // ============================================================

    // OLE2 magic bytes (compound document format)
    private static readonly byte[] Ole2Magic = [0xD0, 0xCF, 0x11, 0xE0];

    private static string ConvertDocToHtml(string filePath)
    {
        try
        {
            var data = File.ReadAllBytes(filePath);

            // Check if it's an OLE2 binary Word document
            if (data.Length > 4 && data[0] == Ole2Magic[0] && data[1] == Ole2Magic[1]
                && data[2] == Ole2Magic[2] && data[3] == Ole2Magic[3])
            {
                return ExtractFromBinaryDoc(data);
            }

            // Otherwise, treat as text-based format (XML Word 2003, MHTML, RTF, etc.)
            return ExtractFromTextDoc(data);
        }
        catch (Exception ex)
        {
            return $"<p>Unable to preview this document: {HttpUtility.HtmlEncode(ex.Message)}</p>";
        }
    }

    /// <summary>
    /// Extract text from a genuine OLE2 binary .doc file.
    /// The actual body text is stored as Latin-1 (cp1252) text interspersed
    /// with binary control data. We extract readable chunks and filter out
    /// embedded XML metadata (SharePoint/OneDrive properties).
    /// </summary>
    private static string ExtractFromBinaryDoc(byte[] data)
    {
        var latin1Text = Encoding.Latin1.GetString(data);

        // Extract all chunks of readable text (20+ printable ASCII chars)
        var allChunks = Regex.Matches(latin1Text, @"[\x20-\x7e\n\r]{20,}")
            .Select(m => m.Value)
            .ToList();

        // Filter OUT chunks that are XML/metadata — these are NOT document content
        var textChunks = allChunks.Where(chunk =>
        {
            var trimmed = chunk.TrimStart();
            // Skip XML declarations, tags, namespaces
            if (trimmed.StartsWith("<")) return false;
            if (trimmed.Contains("xmlns")) return false;
            if (trimmed.Contains("xsd:")) return false;
            if (trimmed.Contains("schemas.openxmlformats")) return false;
            if (trimmed.Contains("schemas.microsoft.com")) return false;
            // Skip OLE/binary metadata keywords
            if (trimmed.StartsWith("Root Entry")) return false;
            if (trimmed.StartsWith("WordDocument")) return false;
            if (trimmed.StartsWith("DocumentSummaryInformation")) return false;
            if (trimmed.StartsWith("SummaryInformation")) return false;
            if (trimmed.StartsWith("Default Paragraph Font")) return false;
            return true;
        })
        .OrderByDescending(c => c.Length)
        .ToList();

        if (textChunks.Count == 0)
            return "<p>Unable to extract readable text from this document.</p>";

        // Combine the text chunks (take the largest ones that are likely body text)
        var sb = new StringBuilder();
        foreach (var chunk in textChunks)
        {
            var cleaned = chunk.Trim();
            if (cleaned.Length < 10) continue;

            // Split on newlines/returns within the chunk
            foreach (var line in cleaned.Split(new[] { '\r', '\n' }, StringSplitOptions.RemoveEmptyEntries))
            {
                var t = line.Trim();
                if (t.Length < 3) continue;
                // Final safety: strip any stray tags
                t = Regex.Replace(t, @"<[^>]+>", " ").Trim();
                if (t.Length < 3) continue;
                sb.AppendLine($"<p>{HttpUtility.HtmlEncode(t)}</p>");
            }
        }

        return sb.Length > 0 ? sb.ToString() : "<p>Document appears to be empty.</p>";
    }

    /// <summary>
    /// Extract text from text-based .doc files (Word 2003 XML, MHTML, RTF).
    /// </summary>
    private static string ExtractFromTextDoc(byte[] data)
    {
        var content = Encoding.UTF8.GetString(data);
        var trimmed = content.TrimStart();

        // Word 2003 XML or HTML-based .doc
        if (trimmed.StartsWith("<?xml") || trimmed.StartsWith("<html", StringComparison.OrdinalIgnoreCase)
            || trimmed.StartsWith("<w:") || trimmed.StartsWith("MIME-Version", StringComparison.OrdinalIgnoreCase))
        {
            // For MHTML, only process the HTML part
            if (trimmed.StartsWith("MIME-Version", StringComparison.OrdinalIgnoreCase))
            {
                var htmlStart = content.IndexOf("<html", StringComparison.OrdinalIgnoreCase);
                if (htmlStart > 0)
                    content = content[htmlStart..];
            }

            // Strip all XML/HTML tags
            var plainText = Regex.Replace(content, @"<style[^>]*>[\s\S]*?</style>", " ", RegexOptions.IgnoreCase);
            plainText = Regex.Replace(plainText, @"<script[^>]*>[\s\S]*?</script>", " ", RegexOptions.IgnoreCase);
            plainText = Regex.Replace(plainText, @"<br\s*/?>", "\n", RegexOptions.IgnoreCase);
            plainText = Regex.Replace(plainText, @"</p>|</div>|</tr>|</li>", "\n", RegexOptions.IgnoreCase);
            plainText = Regex.Replace(plainText, @"<[^>]+>", " ");
            // Decode HTML entities
            plainText = Regex.Replace(plainText, @"&nbsp;", " ");
            plainText = Regex.Replace(plainText, @"&amp;", "&");
            plainText = Regex.Replace(plainText, @"&lt;", "<");
            plainText = Regex.Replace(plainText, @"&gt;", ">");
            plainText = Regex.Replace(plainText, @"&quot;", "\"");
            plainText = Regex.Replace(plainText, @"&#\d+;", " ");
            plainText = Regex.Replace(plainText, @"&\w+;", " ");
            // Collapse whitespace within lines
            plainText = Regex.Replace(plainText, @"[ \t]+", " ");

            if (plainText.Trim().Length < 20)
                return "<p>Unable to extract text from this document.</p>";

            var sb = new StringBuilder();
            foreach (var line in plainText.Split('\n'))
            {
                var t = line.Trim();
                if (t.Length < 2) continue;
                sb.AppendLine($"<p>{HttpUtility.HtmlEncode(t)}</p>");
            }
            return sb.Length > 0 ? sb.ToString() : "<p>Document appears to be empty.</p>";
        }

        // RTF format
        if (trimmed.StartsWith("{\\rtf"))
        {
            var plainText = Regex.Replace(content, @"\\[a-z]+\d*\s?", " ");
            plainText = Regex.Replace(plainText, @"[{}]", "");
            plainText = Regex.Replace(plainText, @"\s+", " ");

            var sb = new StringBuilder();
            foreach (var line in plainText.Split(new[] { "\\par", "\\line" }, StringSplitOptions.RemoveEmptyEntries))
            {
                var t = line.Trim();
                if (t.Length < 2) continue;
                sb.AppendLine($"<p>{HttpUtility.HtmlEncode(t)}</p>");
            }
            return sb.Length > 0 ? sb.ToString() : "<p>Document appears to be empty.</p>";
        }

        return "<p>Unable to preview this document format.</p>";
    }
}
