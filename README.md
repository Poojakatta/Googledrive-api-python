# Googledrive-api-python

**Overview** 

This application allows you to list the files available in your Google Drive, facilitates to read the content inside them and also allows you to query about the data inside the files using OpenAI API(GPT-3.5)

---------------------------------------------------------------------------------------------------------------------------------------------------------------------

To do list:

1. Install all the dependencies with **pip -r requirement.txt**
2. Create an OpenAI account by signing up in www.openai.com and follow the steps from the link(https://www.howtogeek.com/885918/how-to-get-an-openai-api-key/) to get API KEY. Replace the key in code for the variable "openai.api_key"
3. Once you run the program, the application asks you to login to the google account on which you want to perform action. 
4. Give the credentials of the google account and click on 'Continue' incase of any alert message.
5. Upon successful login, it displays the below message 'Go back to terminal.
The authentication flow has completed. You may close this window.'
6. In CLI you will be asked for which action you want to perform.
7. Sample test results are uploaded in tests folder
8. Incase you want to access the spreadsheet other than what is already present in the drive, need to share the spreadsheet to service account mail id mentioned in the service_account.json

