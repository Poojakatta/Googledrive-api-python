import os
import re
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
import gspread
import sys
import openai
import PyPDF2
import io
from googleapiclient.http import MediaIoBaseDownload

# Replace the Openai APIKEY
openai.api_key = "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"
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/documents']

# Generating credentials and connecting to Google drive
flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', scopes=SCOPES)
credentials = flow.run_local_server()
credentials_filename = 'credentials.json'
credentials_data = credentials.to_json()
with open(credentials_filename, 'w') as credentials_file:
    credentials_file.write(credentials_data)
drive_service = build('drive', 'v3', credentials=credentials)


def list_files():
    try:
        # Get all files available in google drive
        result = drive_service.files().list().execute()
        files = result.get('files', [])
        if not files:
            print('No files found')
        else:
            for file in files:
                print(file)
    except Exception as e:
        print("An error occurred while listing all files in the drive:" + str(e))


def read_file(file_id, file_type):
    try:
        #   read the content inside the file based on type of file
        if file_type.lower() == "sheet":
            # to read excel, we need gspread library and share the sheet to service account mail from
            # service_account.json
            gc = gspread.service_account('service_account.json')
            spreadsheet = gc.open_by_key(file_id)
            worksheet = spreadsheet.get_worksheet(0)
            rows = worksheet.get_all_records()
            return str(rows)  # returns all rows from sheet
        elif file_type.lower() == "doc":
            file = drive_service.files().export(fileId=file_id, mimeType='text/plain')
            file_content = file.execute()
            return file_content.decode('utf-8')
        elif file_type.lower() == "pdf":
            return read_pdf_format(file_id)
        else:
            print("Specified file format unavailable!!")
    except Exception as e:
        print("An error occurred while reading" + file_type.lower() + ":" + str(e))


def read_pdf_format(file_id):
    try:
        file = drive_service.files().get_media(fileId=file_id)
        fh = io.FileIO('output.pdf', 'wb')
        downloader = MediaIoBaseDownload(fh, file)
        done = False
        while not done:
            status, done = downloader.next_chunk()
        with open('output.pdf', 'rb') as file:
            # Create a PDF reader object
            pdf_reader = PyPDF2.PdfReader(file)
            num_pages = len(pdf_reader.pages)
            page_text = ""
            # Read the content of each page
            for page_number in range(num_pages):
                page = pdf_reader.pages[page_number]
                page_text += page.extract_text()  # concatenate the data extracted from each page

        return page_text

    except Exception as e:
        print("An error occurred while extracting data from PDF:" + str(e))


def process():
    list_files()
    new_file = True
    while new_file:
        # Google drive api requires file id to get content, so take the file id from input
        file_id = input("Enter the file id of file you want to use:")
        # get file type to perform actions based on type
        file_type = input("Enter the file type (sheet/doc/pdf):")
        # pass the file id and type to get file content
        content = read_file(file_id, file_type)
        if content is not None:
            file_operations = True
            while file_operations:
                action = input("Enter what action you want to perform (Q-Query about details in the file/R-Read):")
                # Query - it takes the file content, passes data to openai api and allows it to read and answer query
                if action.lower() == 'q':
                    try:
                        query = input("Enter the prompt:")
                        # when input data length is greater than 4096,splits content into chunks and performs the action
                        if len(content) > 4096:
                            # remove spaces to reduce the token count
                            cleaned_text = re.sub(r'\s+', '', content)
                            # split the content into chunks
                            chunked_text = [cleaned_text[i:i + 4096] for i in range(0, len(cleaned_text), 4096)]
                            response = ""
                            for chunk in chunked_text:
                                # call gpt api to get response for each chunk and concatenate the response
                                chunk_response = call_gpt(chunk, query)
                                response += chunk_response
                        else:
                            # if the content length is less than 4096 call gpt api directly
                            response = call_gpt(content, query)
                        print(response)

                    except Exception as e:
                        print("An error occurred while Querying Content:" + str(e))
                # when action is to read the content
                elif action.lower() == 'r':
                    try:
                        print(content)
                    except Exception as e:
                        print("An error occurred while reading file:" + str(e))
                else:
                    print(" Choose either Q or R!!")
                    continue

                more_operations = input("Do you want to perform more actions on this file? (Y/N):")
                if more_operations.lower() == 'y':
                    continue
                elif more_operations.lower() == 'n':
                    file_operations = False
                else:
                    print("Invalid Input")
        else:
            print("No content in the selected file!!")
        select_other_file = input("Do you want to choose another file to perform actions? (Y/N):")
        if select_other_file.lower() == 'y':
            new_file = True
        elif select_other_file.lower() == 'n':
            new_file = False
        else:
            print("Invalid Input")


def call_gpt(content, query):
    try:
        # creating prompt to GPT api passing query and the content on which we want to query
        prompt = query + ':' + content
        # using Chatcompletion endpoint and GPT 3.5 turbo model since they gave better response and cost-effective model
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}])

        return response.choices[0].message.content
    except Exception as e:
        print("An error occurred while calling GPT-3.5:" + str(e))


if __name__ == '__main__':
    while 1:
        process()
        # remove the temp file created in run time
        os.remove("output.pdf")
        proceed = input(" Do you want to exit the program? (Y/N):")
        if proceed.lower() == 'n':
            continue
        else:
            sys.exit()
