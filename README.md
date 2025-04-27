# ExcelSlide — PowerPoint Slide Generator from Excel Data
This project automates the creation of multiple PowerPoint slides using Python, pandas, and python-pptx. It pulls structured data from an Excel file and injects it into a predefined PowerPoint template, generating a batch of slides automatically.

Originally built to streamline donor insights reporting, this solution is template-agnostic—it can be easily adapted for any context requiring bulk slide generation from spreadsheet data.

🚀 Features
Batch-generate PowerPoint slides from Excel spreadsheets.

Supports dynamic text replacement in any PowerPoint template.

Reduces manual slide creation by over 20%.

Easily adaptable to different templates and data models.

🛠️ Technologies Used
Python 3.x

pandas (for Excel data handling)

python-pptx (for PowerPoint slide generation)

📦 Installation
Clone the repository:

bash
Copy
Edit
git clone https://github.com/yourusername/ExcelSlide-pptx-slide-generator.git
cd ExcelSlide-pptx-slide-generator
Install dependencies:

bash
Copy
Edit
pip install pandas python-pptx
📈 How to Use
Place your Excel file (e.g., donor_data.xlsx) and your PowerPoint template (e.g., template.pptx) in the project folder.

Update the Python script with the correct Excel file and template paths if necessary.

Run the script:

bash
Copy
Edit
python generate_slides.py
The script will generate a new PowerPoint file with a slide for each entry in the Excel sheet.

🧠 Example Use Cases
Donor insights reports

Sales decks

Project status updates

Event presentations

Any batch PowerPoint generation from tabular data

🧹 Folder Structure
cpp
Copy
Edit
ExcelSlide-pptx-slide-generator/
│
├── generate_slides.py
├── donor_data.xlsx (example)
├── template.pptx (example)
└── README.md
🙌 Contributions
Feel free to fork this project and adapt it to your needs! Pull requests and feedback are welcome.

📬 Contact
Built by André Bernardo — always excited to connect with fellow builders!

✅ Notes:
Replace yourusername with your GitHub username in the git clone link when you create the repo.

Replace filenames if you decide to use different sample files.
