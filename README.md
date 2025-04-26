# Meta Data Clearer

Meta Data Clearer is a Python script designed to remove metadata from various file types, including images, audio files, and Microsoft Office documents (Word, PowerPoint, Excel). It provides a simple graphical interface for selecting a file and clears its metadata.

## Features

- **Supported File Types**:
  - Images: `.jpg`, `.jpeg`, `.png`, `.tiff`, `.bmp`
  - Audio: `.mp3`, `.flac`, `.wav`, `.ogg`, `.m4a`
  - Microsoft Word: `.docx`
  - Microsoft PowerPoint: `.pptx`
  - Microsoft Excel: `.xlsx`
- **Metadata Removal**:
  - Clears metadata such as author, title, subject, comments, keywords, and more.
- **Graphical Interface**:
  - Uses a file selection dialog to choose a file for processing.

## Prerequisites

Before running the script, ensure you have the following Python libraries installed:

- `mutagen`
- `pyexiv2`
- `python-docx`
- `python-pptx`
- `openpyxl`

You can install these libraries using pip:

```bash
pip install mutagen pyexiv2 python-docx python-pptx openpyxl