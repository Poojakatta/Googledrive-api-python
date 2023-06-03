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

openai.api_key = "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"
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/spreadsheets',
          'https://www.googleapis.com/auth/documents']

flow = InstalledAppFlow.from_client_secrets_file('client_secret.json', scopes=SCOPES)
credentials = flow.run_local_server()
credentials_filename = 'credentials.json'
credentials_data = credentials.to_json()
with open(credentials_filename, 'w') as credentials_file:
    credentials_file.write(credentials_data)
drive_service = build('drive', 'v3', credentials=credentials)


def list_files():
    try:
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
        if file_type.lower() == "sheet":
            gc = gspread.service_account('service_account.json')
            spreadsheet = gc.open_by_key(file_id)
            worksheet = spreadsheet.get_worksheet(0)
            rows = worksheet.get_all_records()
            return str(rows)
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
                page_text += page.extract_text()

        return page_text

    except Exception as e:
        print("An error occurred while extracting data from PDF:" + str(e))


def process():
    list_files()
    same_file = True
    while same_file:
        file_id = input("Enter the file id of file you want to use:")
        file_type = input("Enter the file type (sheet/doc/pdf):")
        content = read_file(file_id, file_type)
        if content is not None:
            file_operations = True
            while file_operations:
                action = input("Enter what action you want to perform (Q-Query about details in the file/R-Read):")
                if action.lower() == 'q':
                    try:
                        query = input("Enter the prompt:")
                        if len(content) > 4096:
                            cleaned_text = re.sub(r'\s+', '', content)
                            chunked_text = [cleaned_text[i:i + 4096] for i in range(0, len(cleaned_text), 4096)]
                            response = ""
                            for chunk in chunked_text:
                                chunk_response = call_gpt(chunk, query)
                                response += chunk_response
                        else:
                            response = call_gpt(content, query)
                        print(response)

                    except Exception as e:
                        print("An error occurred while Querying Content:" + str(e))
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

        select_other_file = input("Do you want to choose another file to perform actions? (Y/N):")
        if select_other_file.lower() == 'y':
            same_file = True
        elif select_other_file.lower() == 'n':
            same_file = False
        else:
            print("Invalid Input")


def call_gpt(content, query):
    try:
        prompt = query + ':' + content
        response = openai.ChatCompletion.create(
            model="gpt-3.5-turbo",
            messages=[{"role": "user", "content": prompt}])

        return response.choices[0].message.content
    except Exception as e:
        print("An error occurred while calling GPT-3.5:" + str(e))


if __name__ == '__main__':
    while 1:
        process()
        os.remove("output.pdf")
        proceed = input(" Do you want to exit the program? (Y/N):")
        if proceed.lower() == 'n':
            continue
        else:
            sys.exit()
