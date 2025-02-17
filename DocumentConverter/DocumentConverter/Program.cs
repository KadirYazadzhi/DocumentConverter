using System;
using System.IO;
using System.Text;
using UglyToad.PdfPig;
using Aspose.Words;
using Aspose.Words.Tables;
using Aspose.Words.Drawing;
using System.Linq;

class DocumentConverter {
    static void Main() {
        Console.Title = "Document Converter";
        
        while (true) {
            PrintBanner();
            StartLogic(ReadFilePath());          
        }
    }

    private static void PrintBanner() {
        Console.Clear();
        Console.ForegroundColor = ConsoleColor.DarkCyan;
        Console.WriteLine(@"
██████╗  ██████╗  ██████╗██╗   ██╗███╗   ███╗███████╗███╗   ██╗████████╗     ██████╗ ██████╗ ███╗   ██╗██╗   ██╗███████╗██████╗ ████████╗███████╗██████╗ 
██╔══██╗██╔═══██╗██╔════╝██║   ██║████╗ ████║██╔════╝████╗  ██║╚══██╔══╝    ██╔════╝██╔═══██╗████╗  ██║██║   ██║██╔════╝██╔══██╗╚══██╔══╝██╔════╝██╔══██╗
██║  ██║██║   ██║██║     ██║   ██║██╔████╔██║█████╗  ██╔██╗ ██║   ██║       ██║     ██║   ██║██╔██╗ ██║██║   ██║█████╗  ██████╔╝   ██║   █████╗  ██████╔╝
██║  ██║██║   ██║██║     ██║   ██║██║╚██╔╝██║██╔══╝  ██║╚██╗██║   ██║       ██║     ██║   ██║██║╚██╗██║╚██╗ ██╔╝██╔══╝  ██╔══██╗   ██║   ██╔══╝  ██╔══██╗
██████╔╝╚██████╔╝╚██████╗╚██████╔╝██║ ╚═╝ ██║███████╗██║ ╚████║   ██║       ╚██████╗╚██████╔╝██║ ╚████║ ╚████╔╝ ███████╗██║  ██║   ██║   ███████╗██║  ██║
╚═════╝  ╚═════╝  ╚═════╝ ╚═════╝ ╚═╝     ╚═╝╚══════╝╚═╝  ╚═══╝   ╚═╝        ╚═════╝ ╚═════╝ ╚═╝  ╚═══╝  ╚═══╝  ╚══════╝╚═╝  ╚═╝   ╚═╝   ╚══════╝╚═╝  ╚═╝
                                                                                                                    documentextractor - @kadir_
");
    }

    private static string ReadFilePath() {
        while (true) {
            Console.Write("Enter file path: ");
            string filePath = Console.ReadLine();

            if (filePath.Length == 0) {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("The file path is empty.");
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                
                continue;
            }
            if (!File.Exists(filePath)) {
                Console.ForegroundColor = ConsoleColor.DarkRed;
                Console.WriteLine("The file path doesn't exist.");
                Console.ForegroundColor = ConsoleColor.DarkCyan;
                
                continue;
            }

            return filePath;
        }
    }

    private static void StartLogic(string filePath) {
        
        string outputDir = Path.Combine(Path.GetDirectoryName(filePath), "Extracted");
        Directory.CreateDirectory(outputDir);
        string markdownPath = Path.Combine(outputDir, Path.GetFileNameWithoutExtension(filePath) + ".md");

        try {
            string extension = Path.GetExtension(filePath).ToLower();
            switch (extension) {
                case ".pdf": ExtractFromPdf(filePath, markdownPath, outputDir); break;
                case ".docx":
                case ".odt":
                case ".rtf": ExtractFromWord(filePath, markdownPath, outputDir); break;
                case ".txt": ExtractFromTxt(filePath, markdownPath); break;
                case ".html": ExtractFromHtml(filePath, markdownPath); break;
                default: Console.WriteLine("Unsupported format."); break;
            }
            
            Console.ForegroundColor = ConsoleColor.DarkGreen;
            Console.WriteLine($"✅ Extraction complete! Markdown saved to {markdownPath}");
            Console.ForegroundColor = ConsoleColor.DarkCyan;
        } 
        catch (Exception ex) {
            File.WriteAllText("error.log", $"{DateTime.Now}: {ex.Message}\n{ex.StackTrace}");
            
            Console.ForegroundColor = ConsoleColor.DarkRed;
            Console.WriteLine("❌ Error! Check error.log for details.");
            Console.ForegroundColor = ConsoleColor.DarkCyan;
        }
    }

    static void ExtractFromPdf(string pdfPath, string markdownPath, string outputDir) {
        using var pdf = PdfDocument.Open(pdfPath);
        using StreamWriter writer = new StreamWriter(markdownPath, false, Encoding.UTF8);
        int imageCounter = 1;

        writer.WriteLine("# Extracted PDF Content\n");

        foreach (var page in pdf.GetPages()) {
            writer.WriteLine($"## Page {page.Number}\n");
            writer.WriteLine(page.Text.Trim());
            writer.WriteLine("\n---\n");

            foreach (var image in page.GetImages()) {
                string imageExt = image.RawBytes[0] == 0xFF ? "jpg" : "png"; 
                string imagePath = Path.Combine(outputDir, $"image{imageCounter}.{imageExt}");
                File.WriteAllBytes(imagePath, image.RawBytes.ToArray());
                writer.WriteLine($"![Image {imageCounter}](Extracted/image{imageCounter}.{imageExt})\n");
                imageCounter++;
            }
        }
    }

    static void ExtractFromWord(string wordPath, string markdownPath, string outputDir) {
        Document doc = new Document(wordPath);
        using StreamWriter writer = new StreamWriter(markdownPath, false, Encoding.UTF8);
        int imageCounter = 1;

        writer.WriteLine("# Extracted Word Document\n");

        foreach (Paragraph paragraph in doc.GetChildNodes(NodeType.Paragraph, true)) {
            string text = paragraph.GetText().Trim();
            if (string.IsNullOrEmpty(text)) continue;

            if (paragraph.ParagraphFormat.StyleName.StartsWith("Heading")) {
                int level = paragraph.ParagraphFormat.StyleIdentifier - StyleIdentifier.Heading1 + 1;
                writer.WriteLine($"{new string('#', Math.Clamp(level, 1, 6))} {text}\n");
            } 
            else if (text.StartsWith("- ") || text.StartsWith("* ")) writer.WriteLine(text);
            else if (text.StartsWith("1.") || text.StartsWith("2.")) writer.WriteLine(text);
            else if (text.StartsWith(">")) writer.WriteLine($"> {text}"); 
            else if (paragraph.ParagraphFormat.Style.Font.Bold) writer.WriteLine($"**{text}**");
            else if (paragraph.ParagraphFormat.Style.Font.Italic) writer.WriteLine($"*{text}*");
            else if (paragraph.ParagraphFormat.Style.Font.Underline != Underline.None) writer.WriteLine($"__{text}__");
            else writer.WriteLine($"{text}\n");
        }

        foreach (Table table in doc.GetChildNodes(NodeType.Table, true)) {
            writer.WriteLine(TableToMarkdown(table));
        }

        foreach (Shape shape in doc.GetChildNodes(NodeType.Shape, true)) {
            if (shape.HasImage) {
                string imagePath = Path.Combine(outputDir, $"image{imageCounter}.png");
                shape.ImageData.Save(imagePath);
                writer.WriteLine($"![Image {imageCounter}](Extracted/image{imageCounter}.png)\n");
                imageCounter++;
            }
        }
    }

    static string TableToMarkdown(Table table) {
        StringBuilder sb = new StringBuilder();
        int colCount = table.FirstRow.Cells.Count;

        foreach (Row row in table.Rows) {
            sb.AppendLine("| " + string.Join(" | ", row.Cells.Select(c => c.GetText().Trim())) + " |");
            if (row == table.FirstRow) sb.AppendLine("|" + string.Join("|", Enumerable.Repeat("---", colCount)) + "|");
        }
        return sb.ToString();
    }

    static void ExtractFromTxt(string txtPath, string markdownPath) {
        string text = File.ReadAllText(txtPath, Encoding.UTF8);
        using StreamWriter writer = new StreamWriter(markdownPath, false, Encoding.UTF8);
        writer.WriteLine("# Extracted Text File\n");
        writer.WriteLine(text);
    }

    static void ExtractFromHtml(string htmlPath, string markdownPath) {
        string html = File.ReadAllText(htmlPath, Encoding.UTF8);
        string text = System.Text.RegularExpressions.Regex.Replace(html, "<.*?>", string.Empty);
        using StreamWriter writer = new StreamWriter(markdownPath, false, Encoding.UTF8);
        writer.WriteLine("# Extracted HTML Content\n");
        writer.WriteLine(text);
    }
}
