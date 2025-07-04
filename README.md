Process of leading a team to achieve specific goals within defined constraints like time and budget. 
It involves planning, organizing, and overseeing projects to ensure they are completed successfully, delivering value to stakeholders Project-Management-SW

meu_diretorio/
├── projeto_Obra_Sede.xlsx
├── projeto_DataCenter.xlsx
└── resumo_projetos.xlsx
Set up your email and app password in the send_project_email section before using the sending feature.

To use Gmail, set up an app password to avoid blocks.

The function get_validated_input now accepts a parameter called capitalize to control when to apply .title() on the input.

The "end" command correctly finishes the item registration, without being confused by capitalization.

Requirements to run:

Create a .env file with:
EMAIL_PASSWORD=your_app_password


Install:
pip install python-dotenv openpyxl


App GUI created with Tkinter, including Listbox and buttons for:
  New Project — creates a new project and collects information
  Edit Project — adds items to an existing project
  Delete Project — removes the project's .xlsx file
  Exit — closes the application

Menus at the top
Tabs using ttk.Notebook
Detailed view of items and participants in tables
More organized forms for data entry

How it works:
Main screen: project list on the left, details on the right in tabs.
Project Info tab: shows basic data.
Participants tab: lists participants.
Project Items tab: lists and edits project items.
Menu allows creating a new project or deleting one.
Dialogs to add participants and items with validation.
Automatic email sending to participants.
Prevents duplicate participants and items.
Excel spreadsheets updated and opened automatically.
View of totals and charts in the spreadsheet.
