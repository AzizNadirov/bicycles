# Excel to Word Mapper

A Python application with a graphical user interface for generating multiple Word documents from Excel data using customizable templates.

![Python Version](https://img.shields.io/badge/python-3.6+-blue.svg)
![License](https://img.shields.io/badge/license-MIT-green.svg)

## 🌟 Features

- **User-Friendly Interface**: Simple and intuitive graphical interface for mapping fields and generating documents
- **Smart Field Mapping**: Automatically matches template fields with Excel columns when names match
- **Customizable File Naming**: Generate document names using values from Excel columns
- **Progress Tracking**: Real-time progress bar and status updates during document generation
- **Error Handling**: Robust error checking and document verification
- **Batch Processing**: Generate multiple documents in one go

## 📋 Requirements

```bash
python >= 3.6
pandas
python-docx
docxtpl
```

## 🚀 Installation

1. Clone the repository or download the source code
2. Install required packages:
```bash
pip install pandas python-docx python-docx-template
```

## 📝 Usage

### Creating the Template

1. Create a Word document (.docx) with template fields using the following format:
```
Name: {{name}}
Email: {{email}}
Phone: {{phone}}
```

### Preparing Excel Data

1. Create an Excel file (.xlsx or .xls) with columns that match your template fields
2. Ensure your data is clean and formatted correctly
3. Headers should match template fields for auto-mapping (optional)

### Running the Application

1. Launch the application:
```bash
python excel_to_word_mapper.py
```

2. Using the interface:
   - Click "Browse" to select your Excel file
   - Select your Word template
   - Choose an output directory
   - Map template fields to Excel columns
   - Select columns for filename generation
   - Click "Generate Documents" to start processing

### Filename Pattern

Documents are named using the following pattern:
```
doc-{value1}-{value2}-{n}.docx
```
Where:
- `value1`, `value2`: Values from selected Excel columns
- `n`: Row number (starting from 1)

Example: `doc-John_Smith-2024-01-15-1.docx`

## 🛠️ Example

1. Word Template:
```
Dear {{name}},

Your account balance is {{balance}}.
Your account number is {{account_number}}.

Best regards,
Bank Team
```

2. Excel Data:
```
| name          | balance | account_number |
|---------------|---------|----------------|
| John Smith    | 1000    | ACC001        |
| Jane Doe      | 2000    | ACC002        |
```

3. Generated Files:
- `doc-John_Smith-ACC001-1.docx`
- `doc-Jane_Doe-ACC002-2.docx`

## ⚠️ Important Notes

- Always backup your data before running batch operations
- Template fields must be in the format `{{field_name}}`
- Field names should not contain spaces or special characters
- The application validates generated documents to prevent corruption
- Large files may take longer to process

## 🪲 Troubleshooting

### Common Issues:

1. **No Template Fields Found**
   - Ensure template fields are properly formatted with double curly braces: `{{field_name}}`
   - Check for typos in field names

2. **File Corruption Warning**
   - The application automatically attempts to repair any corrupted files
   - If issues persist, try regenerating the specific document

3. **Missing Mappings**
   - All template fields must be mapped to Excel columns
   - The application will warn you about unmapped fields

## 🤝 Contributing

1. Fork the repository
2. Create a new branch for your feature
3. Submit a pull request with a clear description of your changes

## 📄 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🙏 Acknowledgments

- Built with Python and Tkinter
- Uses python-docx-template for document generation
- Pandas for Excel data handling

## 📬 Contact

For bug reports and feature requests, please use the GitHub issue tracker.

---
Made with ❤️ for document automation