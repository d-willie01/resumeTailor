# Resume Generator Script

## Overview

This Python script generates a professional resume in both DOCX and PDF formats from a structured JSON input file. It uses the `python-docx` library to create a formatted Word document and `docx2pdf` to convert the DOCX into a PDF file automatically.

The resume includes common sections such as:

- Name and Contact Information  
- Education  
- Experience with bullet points  
- Summary  
- Skills  
- Projects (optional)

## Features

- Reads resume data from a JSON file with a predefined structure  
- Formats text with appropriate fonts, sizes, and spacing for readability  
- Generates well-organized sections with bold headings  
- Supports multiple education and experience entries  
- Converts the final DOCX resume into PDF format automatically  
- Cleans up temporary files after conversion  
- Outputs the PDF to a specified folder with a customizable filename  

## Requirements

- Python 3.x  
- `python-docx` (for creating DOCX files)  
- `docx2pdf` (for converting DOCX to PDF)  

You can install dependencies via pip:

```bash
pip install python-docx docx2pdf
