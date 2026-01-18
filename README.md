# PDF to Editable PowerPoint Converter

[![Python Version](https://img.shields.io/badge/python-3.8%2B-blue.svg)](https://www.python.org/downloads/)
[![License](https://img.shields.io/badge/license-MIT-green.svg)](LICENSE)
[![Gemini AI](https://img.shields.io/badge/AI-Gemini%202.5-orange.svg)](https://ai.google.dev/)

A powerful Python tool that converts PDF slides (including image-based PDFs like NotebookLM exports) into **fully editable PowerPoint presentations** using Google Gemini AI for intelligent text extraction.

## Features

- **AI-Powered Extraction**: Uses Google Gemini AI to extract text, styling, and layout from any PDF
- **Visual Preservation**: Keeps the exact look of your PDF as a background image
- **Editable Text Overlay**: Adds editable text boxes on top for easy editing
- **Style Detection**: Extracts font sizes, colors, bold/italic formatting, and text alignment
- **Works with Image-based PDFs**: Perfect for NotebookLM, scanned documents, and design-heavy slides
- **Rate Limit Handling**: Automatic retry with exponential backoff for free tier API usage
- **Model Fallback**: Automatically tries alternative Gemini models if one fails

## Installation

### Prerequisites

- Python 3.8 or higher
- Google AI Studio API key ([Get one free](https://aistudio.google.com/apikey))

### Setup

1. Clone the repository:
```bash
git clone https://github.com/yourusername/pdf-to-pptx.git
cd pdf-to-pptx
```

2. Install dependencies:
```bash
pip install -r requirements.txt
```

3. Set up your Gemini API key:
```bash
# Option 1: Environment variable
export GEMINI_API_KEY='your-api-key-here'

# Option 2: Create a .env file
echo "GEMINI_API_KEY=your-api-key-here" > .env
```

## Quick Start

1. Place your PDF file in the `input/` folder (rename to `slides.pdf` or modify the script)

2. Run the converter:
```bash
python3 converter/pdf_to_pptx.py
```

3. Find your editable PPTX in the `output/` folder

## Configuration

### Environment Variables

| Variable | Description | Required |
|----------|-------------|----------|
| `GEMINI_API_KEY` | Your Google AI Studio API key | Yes |

### Modifying Input/Output Paths

Edit `converter/pdf_to_pptx.py`:
```python
INPUT_PDF = PROJECT_ROOT / "input" / "your-file.pdf"
OUTPUT_PPTX = PROJECT_ROOT / "output" / "your-output.pptx"
```

## How It Works

1. **PDF Rendering**: Each PDF page is rendered as a high-resolution image using PyMuPDF
2. **AI Analysis**: The image is sent to Gemini AI which extracts:
   - All text content with exact wording
   - Text positions (as percentages)
   - Font sizes, colors, and styles
   - Text alignment and bullet levels
3. **PPTX Generation**: For each slide:
   - The PDF page image is added as a background (preserves exact visuals)
   - Editable text boxes are overlaid at detected positions
4. **Fallback**: If AI fails for a slide, it falls back to image-only mode

## Project Structure

```
pdf-to-pptx/
├── converter/
│   └── pdf_to_pptx.py         # Main conversion script with Gemini AI
├── extractor/
│   └── pdf_text_extractor.py  # Legacy text extractor (for text-based PDFs)
├── input/
│   └── slides.pdf             # Place your PDF here
├── output/
│   └── output.pptx            # Generated PowerPoint
├── .env                       # Your API key (create this)
├── .env.example               # Example env file
├── requirements.txt           # Python dependencies
└── README.md
```

## Requirements

```
pdfminer.six==20231228
python-pptx==0.6.23
PyMuPDF==1.24.0
pdf2image==1.17.0
Pillow==10.2.0
google-genai>=1.0.0
```

## Troubleshooting

### Rate Limit Errors (429)

The free tier has limits (10 requests/minute for gemini-2.5-flash-lite). The tool automatically:
- Waits and retries with exponential backoff
- Falls back to alternative models
- Increases delay between requests

If you frequently hit limits:
- Wait a minute and run again
- Or upgrade to a paid API plan

### "GEMINI_API_KEY not set"

Make sure you've set the API key:
```bash
export GEMINI_API_KEY='your-key'
# or create a .env file in the project root
```

### Image-only slides (no editable text)

This happens when AI extraction fails. Check:
- Your API key is valid
- You have API quota remaining
- The PDF image is readable

## License

MIT License - feel free to use, modify, and distribute.

---

Made by [Naseem Ansari](https://www.linkedin.com/in/naseem-ansari-25474b269/)
