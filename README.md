# PDFToDOCXFix

**PDFToDOCXFix** is a Visual Basic for Applications (VBA) macro designed to clean up Word documents that originated from PDF conversions. These types of documents often contain excessive formatting artifacts that make editing — especially on large documents — unstable or painful.

## Features

- Converts all tables into tab-delimited plain text
- Removes:
  - Empty paragraphs
  - Extra line breaks
  - Manual page/section breaks
  - Double spaces
- Normalizes:
  - Font to **Times New Roman, 12pt**
  - Line spacing to **double spaced**
  - Paragraph formatting
  - Document layout to a single column
- Designed to handle **thousands of pages** safely and efficiently

## Why Use This?

When converting large PDF files into editable Word documents, the resulting content can be riddled with:
- Unnecessary tables
- Broken formatting
- Inconsistent spacing
- Invisible anchors and layout glitches

This macro was created in response to user requests dealing with massive documents that **crashed or broke during font changes**. It provides a clean base for formatting, editing, or reflowing text.

## How to Use

### Step 1: Import the Macro
1. Open your Word document.
2. Press `Alt + F11` to launch the **VBA Editor**.
3. Go to **File > Import File...** and select `PDFtoDOCXFix.bas`.
4. Press `F5` or run `PDFtoDOCXFix` manually from the Macros list.

### Step 2: Let It Work
The macro will:
- Clean up layout and structure
- Remove formatting junk
- Display a message box when done

>  **Always make a backup of your document before running any macros.**

## File List

- `PDFtoDOCXFix.bas` — The main VBA macro module
- `README.md` — This file

## License

This project is released under the MIT License.

---

Created with ❤️ for editors working with broken PDF conversions.
