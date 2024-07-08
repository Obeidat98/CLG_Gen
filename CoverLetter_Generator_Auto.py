import datetime
import json
from pathlib import Path
import PySimpleGUI as sg  # Python GUI
from docxtpl import DocxTemplate  # Used to write to the Template
from docx2pdf import convert  # Used to convert from docx to pdf
from bs4 import BeautifulSoup  # Used to find certain elements in the html webpage
import requests  # Used to connect to the LinkedIn Website
from googlesearch import search  # Used to google search the company
import webbrowser  # Used to open the company website link

#region CONFIGURATION

# Function to load configuration data
def load_config():
    if Path(CONFIG_FILE).exists():
        with open(CONFIG_FILE, "r") as file:
            return json.load(file)
    return {}

# Function to save configuration data
def save_config(config):
    with open(CONFIG_FILE, "w") as file:
        json.dump(config, file)

# File path for the configuration file
CONFIG_FILE = "config.json"
config = load_config()  # Load configuration data
template_directory = config.get("TEMPLATE_DIRECTORY", "")
template_path = config.get("TEMPLATE_PATH", "")
output_folder = config.get("OUTPUT_FOLDER", "")

#endregion ---------------------------------------------------------------------------------------------------------

#region GET LINKEDIN DATA

# Function to scrape job details from a LinkedIn job URL
def Get_LinkedIn_Job(job_url):
    # Send a GET request to the job URL
    response = requests.get(job_url)
    while response.status_code != 200:
        response = requests.get(job_url)
        print("Trying to connect to website...")
        if response.status_code == 200:
            break

    if response.status_code == 200:
        print("Connection Established!\n")
        soup = BeautifulSoup(response.text, 'html.parser')

        job_title = soup.find('h1', class_='top-card-layout__title').text.strip()
        company_name = soup.find('a', class_='topcard__org-name-link').text.strip()
        company_fullcity = soup.find('span', class_='topcard__flavor--bullet').text.strip()
        company_city = company_fullcity.split(',')[0]
        hr_name = soup.find('h3', class_='base-main-card__title').text.strip()
        job_id = job_url.split('/')[-2]

        return {
            "Job_Title": job_title,
            "Company_Name": company_name,
            "Company_City": company_city,
            "HR_Name": hr_name,
            "Job_ID": job_id
        }

    else:
        print('Failed to retrieve job details from LinkedIn.')
        return None

#endregion ---------------------------------------------------------------------------------------------------------

#region Get Google Search Results

def search_company_info(company_name, company_city):
    # Formulate the search query
    search_query = f"{company_name} {company_city}"

    # Perform Google search and get search results
    search_results = search(search_query, num_results=5)

    # Return search results as a list
    return list(search_results)

def open_google_search(company_name, company_city):
    # Formulate the search query
    search_query = f"{company_name} {company_city}"

    # Generate the Google search URL
    google_search_url = f"https://www.google.com/search?q={search_query}"

    # Open the Google search results page in the default web browser
    webbrowser.open(google_search_url)

#endregion

#region DOCUMENT

doc = None
if template_path:
    doc = DocxTemplate(template_path)

#endregion ---------------------------------------------------------------------------------------------------------

#region WINDOW & LAYOUT

# Dictionary to store full paths of the templates
template_files_dict = {}

def update_template_dropdown(directory):
    global template_files_dict
    template_files_dict = {str(template.name): str(template) for template in Path(directory).glob("*.docx")}
    return list(template_files_dict.keys())

# Initialize the template files based on the imported folder path from the configuration file
if template_directory:
    template_files = update_template_dropdown(template_directory)
else:
    template_files = []

# Options for skills
skill_options = ["Python Programmierung", "C Programmierung", "Datenanalyse", "Machinelernen", "Softwareentwicklung", "Power Platform", "C/C++", "MATLAB/SIMULINK", "Forschung", "Entwicklung", "Prüfung", "Automatisierung", "Solidworks", "3D CAD"]

layout = [
    [
        sg.Column([

            [sg.Column([
                # Template Directory #
                [sg.Text("Template Directory:", tooltip="Click to select the template directory")],
                # Template Path #
                [sg.Text("Template Path:", tooltip="Select the template path")],
                # Output Folder #
                [sg.Text("Output Folder:", tooltip="Click to select the output folder")],

            ], element_justification='left', justification='left'),

            sg.Column([
                # Template Directory #
                [sg.Input(key="TEMPLATE_DIRECTORY", default_text=template_directory, enable_events=True), sg.FolderBrowse("Browse", target="TEMPLATE_DIRECTORY")],
                # Template Path #
                [sg.Combo(template_files, key="TEMPLATE_PATH", size=(52,1), enable_events=True, readonly=True, default_value=template_path)],
                # Output Folder #
                [sg.Input(key="OUTPUT_FOLDER", default_text=output_folder, enable_events=True), sg.FolderBrowse("Browse", target="OUTPUT_FOLDER")],
                
                ], element_justification='right')],

            # Horizontal Separator
            [sg.HorizontalSeparator()],

            [sg.Column([
                # Personal Gender #
                [sg.Text("Personal Gender:", tooltip="Select the gender of the recipient")],
                # Personal #
                [sg.Text("Personal Name:", tooltip="Enter the name of the recipient")],
                # Company #
                [sg.Text("Company Name:", tooltip="Enter the name of the company")],
                # Address #
                [sg.Text("Company City:", tooltip="Enter the city name")],
                [sg.Text("Street Name:", tooltip="Enter the street name")],
                [sg.Text("Company PLZ:", tooltip="Enter the postal code")],
                # Job #
                [sg.Text("Job Title Long:", tooltip="Enter the long job title")],
                [sg.Text("Job Title Short:", tooltip="Enter the short job title")],
                [sg.Text("Job ID:", tooltip="Enter the Job ID")],
                [sg.Text("Job Skill_1:")],
                [sg.Text("Job Skill_2:")]

            ], element_justification='left', justification='left'),

                sg.Column([
                    # Personal Gender #
                    [sg.Radio('Masculine', "GENDER", default=True, key="MALE", tooltip="Select if the recipient is male"),
                     sg.Radio('Feminine', "GENDER", key="FEMALE", tooltip="Select if the recipient is female"),
                     sg.Radio('Team', "GENDER", key="DIVERSE", tooltip="Select if the recipient's gender is diverse")],
                    # Personal #
                    [sg.Input(key="PERSONAL_NAME")],
                    # Company #
                    [sg.Input(key="COMPANY_NAME")],
                    # Address #
                    [sg.Input(key="COMPANY_CITY")],
                    [sg.Input(key="STREET_NAME")],
                    [sg.Input(key="COMPANY_PLZ")],
                    # Job #
                    [sg.Input(key="JOB_TITLE_LONG")],
                    [sg.Input(key="JOB_TITLE_SHORT")],
                    [sg.Input(key="JOB_ID")],
                    [sg.Combo(skill_options, key="SKILL1", size=(43), readonly=True)],
                    [sg.Combo(skill_options, key="SKILL2", size=(43), readonly=True)],

                ], element_justification='center')],

            # Horizontal Separator
            [sg.HorizontalSeparator()],
            # Buttons #
            [sg.Button("Generate Cover Letter (DOCX)", tooltip="Click to generate cover letter in DOCX format"), sg.Button("Generate Cover Letter (PDF)", tooltip="Click to generate cover letter in PDF format"), sg.Button("Clear Form", tooltip="Click to clear the form"), sg.Exit()],
        ], element_justification='left'),
        sg.VSeperator(),
        sg.Column([
            # LinkedIn Link
            [sg.Text("LinkedIn Job URL:"), sg.Input(key="JOB_URL")],
            [sg.Button("Extract LinkedIn Job Data")],
            # Horizontal Separator
            [sg.HorizontalSeparator()],
            # Search Company Online
            [sg.Button('Search Company'), sg.Button('Search Company (Google View)')],
            [sg.Listbox(values=[], size=(60, 10), key='search_results', enable_events=True)]
        ], vertical_alignment="top")
    ]
]

window = sg.Window("Cover Letter Generator", layout, element_justification="right", resizable=False, progress_bar_color="green")

#endregion ---------------------------------------------------------------------------------------------------------

#region CHECKING LOOP

while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == "Exit":
        break

    # Check which document type is requested
    if event == "Generate Cover Letter (DOCX)":
        generate_docx = True
    elif event == "Generate Cover Letter (PDF)":
        generate_docx = False
    else:
        generate_docx = None

    # Perform all the mutual actions to the DOCX & PDF Generation
    if generate_docx is not None:
        #region DATA PREPARATION
        # Date #
        values["TODAY_DATE"] = datetime.datetime.today().strftime("%d-%m-%Y")

        # Gender #
        if values['MALE']:
            values["PRONOUN"] = "Herr "
            values["GREETING"] = "Sehr geehrter"
        elif values['FEMALE']:
            values["PRONOUN"] = "Frau "
            values["GREETING"] = "Sehr geehrte"
        elif values['DIVERSE']:
            values["PRONOUN"] = ""
            values["GREETING"] = "Sehr geehrte"

        # Personal Name #
        personal_name_parts = values["PERSONAL_NAME"].split()
        if len(personal_name_parts) > 0:
            if (values['DIVERSE']):
                values["LAST_NAME"] = values["PERSONAL_NAME"]
            else:
                values["LAST_NAME"] = personal_name_parts[-1]

        # Date #
        values["TODAY_DATE_SHORT"] = datetime.datetime.today().strftime("%d.%m.%y")
        #endregion

        #region GENERATE COVER LETTER
        if doc is not None and output_folder:
            # Render the template
            doc.render(values)
            output_path = Path(output_folder) / f"{values['COMPANY_NAME']}_{values['JOB_ID']}_Obeidat"
            # Check which file type is requested to be generated
            if generate_docx:
                output_path = output_path.with_suffix(".docx")
            else:
                output_path = output_path.with_suffix(".pdf")

            # Check if the file already exists
            if output_path.exists():
                sg.popup_error("File already exists in the output folder.")
            else:
                # Save the file
                if generate_docx:
                    doc.save(output_path)
                else:
                    output_docx_path = output_path.with_suffix(".docx")
                    doc.save(output_docx_path)
                    convert(output_docx_path, output_path)
                    # Delete the intermediate DOCX file
                    output_docx_path.unlink()
                sg.popup("File saved", f"File has been saved here: {output_path}")
        else:
            sg.popup_error("Please select a template path and output folder.")
        #endregion

    #region CLEAR FORM
    elif event == "Clear Form":
        window["COMPANY_NAME"].update('')
        window["PERSONAL_NAME"].update('')
        window["STREET_NAME"].update('')
        window["COMPANY_PLZ"].update('')
        window["COMPANY_CITY"].update('')
        window["JOB_TITLE_LONG"].update('')
        window["JOB_TITLE_SHORT"].update('')
        window["JOB_ID"].update('')
        window["SKILL1"].update('')
        window["SKILL2"].update('')
        window['MALE'].update(True)
        window['FEMALE'].update(False)
        window['DIVERSE'].update(False)
        window["JOB_URL"].update('')
        window['search_results'].update(values=[])
    #endregion

    #region ONCHANGE_PATH
    elif event == "TEMPLATE_PATH":
        template_filename = values["TEMPLATE_PATH"]
        template_path = template_files_dict.get(template_filename, "")
        if template_path.endswith(".docx"):
            doc = DocxTemplate(template_path)
            print(template_path)
        else:
            sg.popup_error("Please select a .docx file.")

    elif event == "TEMPLATE_DIRECTORY":
        template_directory = values["TEMPLATE_DIRECTORY"]
        template_files = update_template_dropdown(template_directory)
        window["TEMPLATE_PATH"].update(values=template_files)
        print(template_files[0])

    elif event == "OUTPUT_FOLDER":
        output_folder = values["OUTPUT_FOLDER"]
    #endregion

    #region GET_LINKEDIN_DATA
    elif event == "Extract LinkedIn Job Data":
        job_url = values["JOB_URL"]
        if job_url:
            job_details = Get_LinkedIn_Job(job_url)
            if job_details:
                window["COMPANY_NAME"].update(job_details["Company_Name"])
                window["COMPANY_CITY"].update(job_details["Company_City"])
                window["JOB_TITLE_LONG"].update(job_details["Job_Title"])
                window["PERSONAL_NAME"].update(job_details["HR_Name"])
                window["JOB_ID"].update(job_details["Job_ID"])
            else:
                sg.popup_error("Failed to extract job details from LinkedIn.")
        else:
            sg.popup_error("Please enter a LinkedIn job URL.")
    #endregion

    #region Search Company via Google
    elif event == 'Search Company':
        job_details = {
            "Company_Name": values["COMPANY_NAME"],
            "Company_City": values["COMPANY_CITY"]
        }
        search_results = search_company_info(job_details["Company_Name"], job_details["Company_City"])
        window['search_results'].update(values=search_results)
    elif event == 'Search Company (Google View)':
        job_details = {
            "Company_Name": values["COMPANY_NAME"],
            "Company_City": values["COMPANY_CITY"]
        }
        open_google_search(job_details["Company_Name"], job_details["Company_City"])
    elif event == 'search_results':
        selected_item = values['search_results'][0]
        webbrowser.open(selected_item)
    #endregion

#endregion ---------------------------------------------------------------------------------------------------------

#region CONFIGURATION

# Save the last selected template directory, template path, and output folder to the configuration file
config["TEMPLATE_DIRECTORY"] = template_directory
config["TEMPLATE_PATH"] = template_path
config["OUTPUT_FOLDER"] = output_folder
save_config(config)

#endregion ---------------------------------------------------------------------------------------------------------

window.close()
