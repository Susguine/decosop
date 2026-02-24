using System.Text;
using System.Text.RegularExpressions;
using System.Web;
using UglyToad.PdfPig;

namespace DecoSOP.Services;

/// <summary>
/// Extracts HTML body content from files for use in Web SOPs and Web Docs.
/// Used during batch import to auto-generate web-viewable content.
/// </summary>
public static class WebContentExtractor
{
    /// <summary>
    /// Extract HTML body content from the file at the given path.
    /// Returns an HTML fragment suitable for inline rendering.
    /// Never returns null; always returns at least a title-only placeholder.
    /// </summary>
    public static string ExtractHtml(string filePath, string title)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        var safeTitle = HttpUtility.HtmlEncode(title);

        try
        {
            return ext switch
            {
                ".docx" or ".xlsx" or ".doc" => ExtractOfficeHtml(filePath, safeTitle),
                ".pdf" => ExtractPdfHtml(filePath, safeTitle),
                ".txt" or ".csv" => ExtractTextHtml(filePath, safeTitle),
                ".rtf" => ExtractRtfHtml(filePath, safeTitle),
                ".pptx" => ExtractPptxHtml(filePath, safeTitle),
                ".png" or ".jpg" or ".jpeg" or ".gif" or ".bmp" or ".webp" => ExtractImageHtml(filePath, safeTitle),
                _ => $"<h1>{safeTitle}</h1>\n<p><em>This file is available in the file viewer.</em></p>"
            };
        }
        catch
        {
            return $"<h1>{safeTitle}</h1>\n<p><em>Content could not be extracted from this file.</em></p>";
        }
    }

    private static string ExtractOfficeHtml(string filePath, string safeTitle)
    {
        var bodyHtml = DocumentPreviewService.ExtractBodyHtml(filePath);
        if (!string.IsNullOrWhiteSpace(bodyHtml))
            return $"<h1>{safeTitle}</h1>\n{bodyHtml}";
        return $"<h1>{safeTitle}</h1>\n<p><em>Content could not be extracted from this document.</em></p>";
    }

    private static string ExtractPdfHtml(string filePath, string safeTitle)
    {
        using var document = PdfDocument.Open(filePath);
        var sb = new StringBuilder();
        sb.AppendLine($"<h1>{safeTitle}</h1>");

        foreach (var page in document.GetPages())
        {
            var text = page.Text;
            if (string.IsNullOrWhiteSpace(text)) continue;

            var lines = text.Split('\n');
            var currentPara = new StringBuilder();
            foreach (var line in lines)
            {
                var trimmed = line.Trim();
                if (string.IsNullOrEmpty(trimmed))
                {
                    if (currentPara.Length > 0)
                    {
                        sb.AppendLine($"<p>{HttpUtility.HtmlEncode(currentPara.ToString().Trim())}</p>");
                        currentPara.Clear();
                    }
                }
                else
                {
                    if (currentPara.Length > 0)
                        currentPara.Append(' ');
                    currentPara.Append(trimmed);
                }
            }
            if (currentPara.Length > 0)
                sb.AppendLine($"<p>{HttpUtility.HtmlEncode(currentPara.ToString().Trim())}</p>");
        }

        var result = sb.ToString();
        if (result.Length <= safeTitle.Length + 30)
            return $"<h1>{safeTitle}</h1>\n<p><em>No readable text could be extracted from this PDF.</em></p>";

        return result;
    }

    private static string ExtractTextHtml(string filePath, string safeTitle)
    {
        var text = File.ReadAllText(filePath);
        if (string.IsNullOrWhiteSpace(text))
            return $"<h1>{safeTitle}</h1>\n<p><em>File is empty.</em></p>";

        if (text.Length > 100_000)
            text = text[..100_000] + "\n\n[... truncated ...]";

        return $"<h1>{safeTitle}</h1>\n<pre>{HttpUtility.HtmlEncode(text)}</pre>";
    }

    private static string ExtractRtfHtml(string filePath, string safeTitle)
    {
        var content = File.ReadAllText(filePath);
        var plainText = Regex.Replace(content, @"\\[a-z]+\d*\s?", " ");
        plainText = Regex.Replace(plainText, @"[{}]", "");
        plainText = Regex.Replace(plainText, @"\s+", " ").Trim();

        if (string.IsNullOrWhiteSpace(plainText))
            return $"<h1>{safeTitle}</h1>\n<p><em>No readable text found.</em></p>";

        var sb = new StringBuilder();
        sb.AppendLine($"<h1>{safeTitle}</h1>");
        foreach (var segment in plainText.Split(["\\par", "\\line"], StringSplitOptions.RemoveEmptyEntries))
        {
            var t = segment.Trim();
            if (t.Length < 2) continue;
            sb.AppendLine($"<p>{HttpUtility.HtmlEncode(t)}</p>");
        }
        return sb.ToString();
    }

    private static string ExtractPptxHtml(string filePath, string safeTitle)
    {
        using var pptDoc = DocumentFormat.OpenXml.Packaging.PresentationDocument.Open(filePath, false);
        var presentationPart = pptDoc.PresentationPart;
        if (presentationPart?.Presentation?.SlideIdList is null)
            return $"<h1>{safeTitle}</h1>\n<p><em>No slides found.</em></p>";

        var sb = new StringBuilder();
        sb.AppendLine($"<h1>{safeTitle}</h1>");

        int slideNum = 0;
        foreach (var slideId in presentationPart.Presentation.SlideIdList
            .Elements<DocumentFormat.OpenXml.Presentation.SlideId>())
        {
            slideNum++;
            var slidePart = (DocumentFormat.OpenXml.Packaging.SlidePart?)presentationPart
                .GetPartById(slideId.RelationshipId!);
            if (slidePart?.Slide is null) continue;

            var slideText = slidePart.Slide.InnerText?.Trim();
            if (!string.IsNullOrWhiteSpace(slideText))
            {
                sb.AppendLine($"<h2>Slide {slideNum}</h2>");
                sb.AppendLine($"<p>{HttpUtility.HtmlEncode(slideText)}</p>");
            }
        }

        return sb.Length > safeTitle.Length + 30
            ? sb.ToString()
            : $"<h1>{safeTitle}</h1>\n<p><em>No readable text found in slides.</em></p>";
    }

    private static string ExtractImageHtml(string filePath, string safeTitle)
    {
        var ext = Path.GetExtension(filePath).ToLowerInvariant();
        var mimeType = ext switch
        {
            ".png" => "image/png",
            ".jpg" or ".jpeg" => "image/jpeg",
            ".gif" => "image/gif",
            ".bmp" => "image/bmp",
            ".webp" => "image/webp",
            _ => "image/png"
        };

        var bytes = File.ReadAllBytes(filePath);
        var base64 = Convert.ToBase64String(bytes);
        return $"<h1>{safeTitle}</h1>\n<img src=\"data:{mimeType};base64,{base64}\" alt=\"{safeTitle}\" style=\"max-width: 100%; height: auto;\" />";
    }
}
