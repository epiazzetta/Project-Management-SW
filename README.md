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
