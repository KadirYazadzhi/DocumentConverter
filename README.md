# 📜 Document Converter

## 🚀 Overview
The **Document Converter** is a powerful console-based tool designed to **extract text and images** from various document formats and convert them into **Markdown** (.md) files. This tool supports multiple formats, including:

✅ **PDF** 📄
✅ **Word Documents** (.docx, .odt, .rtf) 📝
✅ **Plain Text** (.txt) 📃
✅ **HTML** (.html) 🌐

The extracted content is neatly formatted, making it easy to use for documentation, note-taking, and further processing.


## 🛠 Features
- 🔍 **Automatic document detection** – Reads the file extension and processes accordingly.
- 🖼 **Extracts images** – Saves images from PDF and Word documents.
- 🏷 **Preserves formatting** – Headings, lists, bold/italic text are retained.
- 📂 **Creates an 'Extracted' folder** – Stores output markdown and images.
- ⚡ **Handles errors gracefully** – Logs errors to `error.log` for debugging.
- 🎨 **Stylish console banner** – Adds a nice touch to the CLI experience.


## 📥 Installation & Requirements
### 🖥 Prerequisites
Ensure you have the following installed:
- **.NET Framework / .NET Core**
- **Aspose.Words** (for Word document parsing)
- **UglyToad.PdfPig** (for PDF extraction)

To install dependencies:
```sh
# Install Aspose.Words (NuGet package)
dotnet add package Aspose.Words

# Install UglyToad.PdfPig (NuGet package)
dotnet add package UglyToad.PdfPig
```

## 🎬 How to Use
1️⃣ **Run the application**
```sh
dotnet run
```

2️⃣ **Enter the file path** when prompted.
3️⃣ The program automatically detects the format and extracts the content.
4️⃣ The extracted **Markdown file** and **images** are saved in the `Extracted/` directory.


## 📂 Output Structure
```
📂 Extracted/
 ├── 📜 extracted_document.md  # Markdown file with extracted text
 ├── 🖼 image1.png              # Extracted images (if any)
 ├── 🖼 image2.jpg              # Additional images
```


## 🧩 Supported Formats & Extraction Logic
### 📝 **Word Documents (.docx, .odt, .rtf)**
- ✅ Extracts **headings**, **bold**, **italic**, and **underline** formatting.
- ✅ Retains **lists** and **quotes**.
- ✅ Extracts **tables** and converts them into Markdown.
- ✅ Saves **images** inside the `Extracted/` folder.

### 📄 **PDF Files (.pdf)**
- ✅ Extracts **text page by page**.
- ✅ Saves **images** from the PDF into `Extracted/`.

### 📃 **Plain Text (.txt)**
- ✅ Copies text **as-is** into a Markdown file.

### 🌐 **HTML Files (.html)**
- ✅ Converts HTML content into Markdown-friendly format.


## 🏗 Code Breakdown
### 🔄 **Main Execution Flow**
The program follows this simple **looped workflow**:
1. **Display Banner** 🎨 – Shows a stylized title screen.
2. **Prompt for File Path** 📂 – Requests user input.
3. **Detect File Type** 🔍 – Determines the format based on the extension.
4. **Extract Content** 📜 – Calls the appropriate extraction function.
5. **Save Output** 💾 – Stores results in `Extracted/`.
6. **Handle Errors** ❌ – Logs issues in `error.log`.

### 🔑 **Key Functions**
| Function | Description |
|----------|-------------|
| `PrintBanner()` | Displays the ASCII banner. |
| `ReadFilePath()` | Prompts user to enter a file path and validates it. |
| `StartLogic(filePath)` | Determines file type and calls appropriate extraction function. |
| `ExtractFromPdf(filePath, markdownPath, outputDir)` | Extracts text and images from PDF. |
| `ExtractFromWord(filePath, markdownPath, outputDir)` | Extracts formatted text, tables, and images from Word documents. |
| `ExtractFromTxt(filePath, markdownPath)` | Converts plain text into Markdown. |
| `ExtractFromHtml(filePath, markdownPath)` | Extracts HTML content into Markdown. |


## ⚠️ Error Handling
If an error occurs during execution:
- The error details (message + stack trace) are logged to `error.log`.
- A ❌ message is displayed on the console.
- The program continues running without crashing.


## 📌 Future Improvements
🔹 Add support for **Excel (.xlsx, .csv)** conversion 📊  
🔹 Improve **image recognition** and metadata extraction 🖼  
🔹 Implement **batch processing** for multiple files 📁  


## 🤝 Contributing
Want to improve this tool? Feel free to fork, modify, and submit a pull request!

```sh
git clone https://github.com/yourusername/document-converter.git
cd document-converter
dotnet run
```


## 📜 License
This project is licensed under the **MIT License** - see the LICENSE file for details.

### ⚖️ What does this mean?
- ✅ You are free to use, modify, and distribute this software.
- ✅ You can use it for both personal and commercial projects.
- ❌ You cannot hold the author liable for any damages or misuse.
  
---

🌟 If you like this project, consider giving it a star! ⭐

