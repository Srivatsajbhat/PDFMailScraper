# PDFMailScraper

PDFMailScraper is a Python tool that automates the process of extracting email addresses from PDF files and storing them in a Word document. This tool can be particularly useful when you have a collection of PDF files containing contact information and you want to quickly compile a list of email addresses.

## Features

- Extracts email addresses from multiple PDF files within a specified directory.
- Avoids duplication by automatically filtering out duplicate email addresses.
- Saves the extracted email addresses in a Word document for easy access and reference.
- Customizable to fit specific project needs.

## Prerequisites

Before using PDFMailScraper, ensure you have the following installed:

- Python 3.x
- Required Python packages: PyMuPDF (for PDF processing), python-docx (for Word document manipulation)

You can install the required packages using the following command:
```
pip install PyMuPDF python-docx
```

## Usage

1. Clone this repository or download the ZIP archive.
2. Navigate to the project directory in your terminal.

To run the tool, execute the following command:

```
python mail_scraper.py
```
**After running this command you will be asked to enter the path of the directory where you saved all your PDFs**



The tool will process each PDF file in the specified directory, extract unique email addresses, and store them in a Word document.

## Contributing

Contributions to PDFMailScraper are welcome! If you find any issues or have suggestions for improvements, feel free to open an issue or submit a pull request.


---

**Note**: While this tool strives to accurately extract email addresses, results may vary based on PDF formatting and content. Please review the extracted email addresses for accuracy.



