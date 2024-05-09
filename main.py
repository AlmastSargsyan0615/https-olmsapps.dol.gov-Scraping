import requests
import json
import csv
from datetime import datetime, timedelta
import os
import requests
import pdfkit
import shutil
import time

filename = "company_reports_" + datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + ".csv"
foldername = "company_reports_" + datetime.now().strftime("%Y-%m-%d_%H-%M-%S") + "_folder"
path_wkthmltopdf = r"C:\Program Files\wkhtmltopdf\bin\wkhtmltopdf.exe" 
config = pdfkit.configuration(wkhtmltopdf=path_wkthmltopdf)

def make_pdf_from_url(url, output_filename):
    def get_html_content(url):
        response = requests.get(url)
        content_type = response.headers.get('Content-Type', '')

        # Check if the Content-Type header contains 'text/html'
        is_html = 'text/html' in content_type.lower()

        # If it's HTML, return True and the HTML content
        if is_html:
            return True, response.content.decode('utf-8')  # Decode the content to a string
        else:
            return False, response.content

    is_html, content = get_html_content(url)
    if is_html:
        print("The response is an HTML page.")
        pdfkit.from_url(url, output_filename, configuration=config)
        return output_filename
    else:
        print("The response is not an HTML page.")
        with open(output_filename, 'wb') as file:
            file.write(content)
        return output_filename

def move_file_to_destination(source_relative_path, destination_relative_path):
    # Get the directory of the current script
    script_directory = os.path.dirname(os.path.abspath(__file__))

    # Get the absolute paths by joining with the script directory
    source_absolute_path = os.path.join(script_directory, source_relative_path)
    destination_absolute_path = os.path.join(script_directory, destination_relative_path)

    # Check if the source file exists
    if os.path.exists(source_absolute_path):
        # Check if the destination directory exists, if not create it
        if not os.path.exists(destination_absolute_path):
            os.makedirs(destination_absolute_path)
        
        # Get the filename
        filename = os.path.basename(source_absolute_path)

        # Create the full destination path
        destination_full_path = os.path.join(destination_absolute_path, filename)

        # Check if the file already exists in the destination directory
        if os.path.exists(destination_full_path):
            print(f"File '{filename}' already exists in the destination directory. Skipping move.")
        else:
            # Move the file to the destination directory
            shutil.move(source_absolute_path, destination_full_path)
            print("File move initiated.")
            
            # Wait until the move operation is completed
            while not os.path.exists(destination_full_path):
                time.sleep(0.1)  # Adjust sleep time as needed
            
            print("File moved successfully.")
    else:
        print("Source file does not exist.")

def replace_invalid_characters(string):
    invalid_chars = '<>|:\/\"*?'
    for char in invalid_chars:
        string = string.replace(char, '-')
    return string

def write_company_reports(data):
    with open(filename, 'a', newline='') as csvfile:  
        fieldnames = ['File Number', 'Employer Name', 'City', 'State', 'Report',
                      'Received Date', 'Year', 'Receiver Organization', 'Report Url', 'Filer Type']
        writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
        
        # Check if the file is empty, if so, write the header
        if csvfile.tell() == 0:
            writer.writeheader()
        

        for report in data['reports']:
            writer.writerow({
                'File Number': data['fileNumber'],
                'Employer Name': data['employerName'],
                'City': data['employerCity'],
                'State': data['employerState'],
                'Report': report['reportName'],
                'Received Date': report['receivedDate'],
                'Year': report['yrCovered'],
                'Receiver Organization': report['receiverOrganization'],
                'Report Url': report['reportUrl'],
                'Filer Type': data['filerType']
            })

            employerName_folderpath = replace_invalid_characters(data['employerName'])
            destination_relative_path = f"{foldername}/{employerName_folderpath}_{data['fileNumber']}_{data['filerType']}"

            move_file_to_destination( make_pdf_from_url(report['reportUrl'], report['reportName']), destination_relative_path)
   
def get_company_reports(sr_num):
    url = "https://olmsapps.dol.gov/olpdr/GetLM10FilerDetailServlet"
    payload = {
        "srNum": "E-" + sr_num
    }

    response = requests.post(url, data=payload)
    report_cnt = 1
    if response.status_code == 200:
        response_data = json.loads(response.text)
        detail_list = response_data.get("detail", [])

        reports = []
        for detail in detail_list:
            report_name = "E-" + str(sr_num) + "_" + str(report_cnt) + ".pdf"

            receive_date = detail.get("receiveDate")
            if receive_date:
                # Convert receiveDate to string format (MM/DD/YYYY) if it exists
                receive_date = (datetime.fromtimestamp(receive_date / 1000) + timedelta(days=1)).strftime("%m/%d/%Y")
            else:
                receive_date = ""

            # Concatenate subLabOrg1 and subLabOrg2 into one string
            sub_lab_orgs = "{} {}".format(detail.get("subLabOrg1", ""), detail.get("subLabOrg2", ""))

            rptId = detail.get("rptId")
            formLink = detail.get("formLink")
            reportUrl = f"https://olmsapps.dol.gov/query/orgReport.do?rptId={rptId}&rptForm={formLink}"
            report = {
                'reportName': report_name,
                'receivedDate': receive_date,
                'yrCovered': detail.get("yrCovered"),
                'receiverOrganization': sub_lab_orgs.strip(),
                'reportUrl': reportUrl
            }
            reports.append(report)
            report_cnt = report_cnt + 1

        return reports
    else:
        print("Error:", response.status_code)
        return None

def extract_companies(response_text):
    # Parse JSON response
    response_data = json.loads(response_text)

    # Extract company information
    company_info_list = response_data.get("filerList", [])

    # Creating new list of companies
    for company_info in company_info_list:
        company = {
            'fileNumber': "E-" + str(company_info['srNum']),
            'employerName': company_info.get("companyName"),
            'employerCity': company_info.get("companyCity"),
            'employerState': company_info.get("companyState"),
            "srNum": company_info.get("srNum"),
            "filerType": company_info.get("filerType"),
            "reports" : get_company_reports(str(company_info.get("srNum")))
        }
        print("------------------------------------------------------------------------------------")
        print(company)
        write_company_reports(company)
        print("======================================================================================")
    return len(company_info_list)

pageNum = 1
while True:
    url = "https://olmsapps.dol.gov/olpdr/GetLM10FilerListServlet"
    payload = {"clearCache": "F", "page": str(pageNum)}
    response = requests.post(url, data=payload)

    if response.status_code == 200:
        response_text = response.text
    else:
        print("Error:", response.status_code)

    companies_onepage = extract_companies(response_text)

    if companies_onepage == 0:
        break
    else:
        pageNum += 1
