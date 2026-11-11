Convert Tools
A Windows desktop utility for batch-converting Microsoft Word documents
to PDF and combining multiple PDF files into a single document.

Convert Tools is designed to make repetitive document preparation
easier, especially when handling many files at once.
Features
Word to PDF
- Convert multiple .doc and .docx files in one batch.
- Select individual files instead of having to process an entire
  folder.
- Drag and drop Word files into the application.
- Choose an output folder for converted PDFs.
- Preserve the original file names.
- Display conversion progress and an estimated time remaining.
- Retry failed conversions (up to 3 attempts).
- Resume an interrupted batch using the saved conversion state.
- Process files in batches of up to 100.
- Generate an Excel conversion report.
- Detect and report files that cannot be converted.
PDF Combiner
- Select multiple PDF files independently from the Word conversion
  list.
- Drag and drop PDF files into the application.
- Review the selected files before combining.
- Merge PDFs in the order in which they are selected.
- Save the merged document as combined_all.pdf.
Supporting files
- conversion_report.xlsx --- conversion results and status report.
- conversion_state.sqlite3 --- local progress/checkpoint data used
  to resume a batch.
Requirements
- Windows 10 or Windows 11.
- Python 3.9 or later for running from source (use a Python version
  supported by the dependencies).
- Microsoft Word desktop application installed and activated for
  .doc / .docx conversion.
- Internet access during the initial dependency installation.
Important: Word-to-PDF conversion uses Microsoft Word automation
(COM). Microsoft Word desktop must be installed on the computer
running the application. The standalone executable does not include
Microsoft Word.

Run from Source
1. Clone this repository:
   git clone https://github.com/<your-username>/<your-repository>.git
   cd <your-repository>
2. Create and activate a virtual environment:
   py -m venv .venv
   .venv\Scripts\activate
3. Install dependencies:
   python -m pip install --upgrade pip
   pip install -r requirements.txt
4. Run the application:
   python convert_tools.py
Build a Windows Executable
The repository includes a Windows build script for creating a standalone
executable with PyInstaller.
1. Open Command Prompt in the project directory.
2. Run:
   BUILD_RELEASE.bat
3. Wait for the build to finish.
4. Follow the prompts in the packaging step to create the
   password-protected source archive and distribution package.
The resulting release files are generated in the release/output
directory configured by the build scripts.
Build the executable on Windows. A build performed on Linux or macOS
will not produce a native Windows executable.

Typical Distribution
The packaged application is intended to be distributed as a ZIP archive
containing:
Convert_Tools/
├── Convert Tools.exe
├── Source_Code_Locked.zip
└── README.txt
End users can run Convert Tools.exe without installing Python. They
still need Microsoft Word desktop installed for Word document
conversion.
Project Structure
The exact files may vary by release. The main project files include:
.
├── convert_tools.py
├── requirements.txt
├── BUILD_RELEASE.bat
├── package_release.py
└── README.md
Notes and Limitations
- Conversion progress is measured by completed documents. Microsoft
  Word automation does not provide a reliable internal percentage for
  the conversion of an individual document.
- Password-protecting the source archive helps restrict casual access
  to the packaged source code; it does not make the executable
  impossible to inspect or reverse engineer.
- Do not put passwords, API keys, personal data, or other secrets in
  the source code or executable.
- Always review the generated PDFs and conversion report, especially
  when processing important business documents.
- Keep a backup of original documents. The application should not be
  treated as a replacement for document backup.
Troubleshooting
The application cannot convert a Word file - Confirm that Microsoft
Word desktop is installed and activated. - Open the document manually in
Word to check whether it is corrupted or prompts for repair. - Close any
dialog boxes left open by Microsoft Word and try again.

A file is reported as missing - Check that the source file still
exists and that the application has permission to read it. - Confirm
that the selected output folder is writable.

The build fails - Run BUILD_RELEASE.bat from Command Prompt so the
full error output remains visible. - Check that Python is installed and
available through the py launcher. - Retry after confirming that
dependencies can be installed from requirements.txt.
Contributing
Issues and suggestions are welcome. Before submitting a pull request,
please describe the problem or proposed change and include clear
reproduction steps where applicable.
License
license by me. Unless a license is added to this
repository, all rights are reserved by the copyright holder. Do not
assume that the source code may be reused, modified, or redistributed.
