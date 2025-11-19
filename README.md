# QLD Building Forms Application

A Python3 + tkinter application to assist engineers with filling out Form 12 inspection forms.

## Features

- GUI for filling out inspection forms with configurable default values
- Save form data to JSON files
- Load form data from JSON files
- Generate DOCX files with filled form data from template.docx
- Automatic creation of JSON file with same name as DOCX
- Global storage of 'building certifier' and 'appointed competent person' details
- Unique name enforcement for appointed competent persons with automatic override
- + button to select from previously entered details
- Manual signature handling (security conscious approach)
- Command-line flag support for alternate templates
- Multiline text areas for longer form entries
- Mouse wheel/trackpad scrolling support
- All appointed competent person fields stored in global.json:
  * Name
  * Company
  * Contact person
  * Business phone
  * Mobile
  * Email
  * Postal address
  * Suburb/locality (postal)
  * State (postal)
  * Postcode (postal)
  * Licence type
  * Licence number

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

1. Run the application with default template:

```bash
python3 src/main.py
```

2. Run with a custom template:

```bash
python3 src/main.py --template /path/to/your/template.docx
```

3. Or use the run script:

```bash
python3 run_app.py
```

4. Fill out the form fields or load a previously saved form
5. Use the "Save" button to save form data to a JSON file
6. Use the "Generate DOCX" button to create a populated DOCX
7. Use the "Load" button to load a previously saved form
8. Use the "Reset" button to clear all form fields
9. For building certifier and competent person fields, use the "+" button to select from previously entered details

## Configuration

- Default values can be configured in `defaults.json`
- Global details for building certifier and appointed competent person are stored in `global.json`
- Appointed competent person entries are automatically stored/updated when saving or generating forms
- All 12 appointed competent person fields are preserved in global.json with unique name enforcement

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
- `template.docx`: Template for DOCX generation
- `requirements.txt`: Python dependencies
- `run_app.py`: Convenient start script
- `app.command`: macOS application launcher (ignored by git)
- `Form12 Inspections/`: Directory for local inspection files (ignored by git)