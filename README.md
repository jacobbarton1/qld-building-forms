# Inspection Form Application

A Python3 + tkinter application to assist engineers with filling out inspection forms.

## Features

- GUI for filling out inspection forms with configurable default values
- Save form data to JSON files
- Load form data from JSON files
- Generate DOCX files with filled form data from template.docx
- Automatic creation of JSON file with same name as DOCX
- Global storage of 'building certifier' and 'appointed competent person' details
- Duplicate prevention for global details
- + button to select from previously entered details

## Requirements

- Python 3.7+
- python-docx
- tkinter (usually included with Python)

## Installation

1. Clone this repository
2. Install the required packages:

```bash
pip install -r requirements.txt
```

3. Place your `template.docx` in the root directory

## Usage

1. Run the application:

```bash
python src/main.py
```

2. Fill out the form fields or load a previously saved form
3. Use the "Save" button to save form data to a JSON file
4. Use the "Generate PDF" button to create a populated PDF
5. Use the "Load" button to load a previously saved form
6. Use the "Reset" button to clear all form fields
7. For building certifier and competent person fields, use the "+" button to select from previously entered details

## Configuration

- Default values can be configured in `defaults.json`
- Global details for building certifier and appointed competent person are stored in `global.json`

## Form Fields

The application includes fields for:

- Project information
- Location and inspection details
- Structural elements
- Services (electrical, plumbing, mechanical)
- Safety issues
- Compliance status
- Recommendations
- Building certifier details
- Appointed competent person details

## File Structure

- `src/main.py`: Main application code
- `defaults.json`: Default values for form fields
- `global.json`: Global details for building certifier and competent person
- `template.pdf`: Template for PDF generation (first 2 pages used)
- `requirements.txt`: Python dependencies