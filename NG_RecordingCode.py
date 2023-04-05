import requests
import datetime
import openpyxl
import pandas as pd
import concurrent.futures  # Import the concurrent.futures library for parallel processing
import shutil
from urllib.parse import urlparse

shutil.os.makedirs("C:\\AlphaData\\Recordings\\NextGen\\April\\3")
# Create a function to download the recordings from the URL links
def download_recording(url, index):
    try:
        if pd.isna(url):  # Check if the url is empty
            skip = "skip"
            print(f"Skipping {url}")
            return None
        
        parsed_url = urlparse(url)
        if not parsed_url.scheme:
            url = "https://" + url

        response = requests.get(url, timeout=10, verify=False)
  # increase the timeout to 10 seconds
        if response.status_code == 200:
            filename = "C:\\AlphaData\\Recordings\\NextGen\\April\\3\\" + str(df.iloc[index]['Medicare#']).replace('\t', '').replace('\n', '').replace('::::','').replace('::','').replace(':','').replace('/','').replace(' ', '') + '.mp3'
            with open(filename, 'wb') as f:
                f.write(response.content)
            print(f"Successfully downloaded {url}")
            return filename
        else:
            # If the request is not successful, return None
            print(f"Error downloading {url}")
            return None
    except requests.exceptions.Timeout:
        # If the download times out, return None
        print(f"Timeout error downloading {url}")
        return None
    except requests.exceptions.ConnectionError:
        print(f"Connection error downloading {url}")
        return None



# Create a new workbook and worksheet
workbook = openpyxl.Workbook()
worksheet = workbook.active
worksheet.title = 'NG_04-03-2023_ALPHA'

# Add headers to the worksheet
headers = ['Timestamp', 'Last Name', 'First Name', 'Gender', 'DOB', 'MBI Number', 'Address', 'Customer City', 'State', 'Zip', 'Contact Number', 'Pverify', 'Recordings Links', 'ALPHA_Timestamp', 'Downloaded?', 'Error Comment:', 'FirstName Matched', 'LastName Matched', 'DOB Matched', 'Address Matched', 'MBI Number Matched', 'Gave Consent?']
worksheet.append(headers)

# Save the workbook
workbook.save('NG_04-03-2023_ALPHA.xlsx')

# Open the workbook again in append mode
workbook = openpyxl.load_workbook('NG_04-03-2023_ALPHA.xlsx')
worksheet = workbook.active

for column in worksheet.columns:
    max_length = 0
    column_name = column[0].column_letter
    for cell in column:
        try:
            if len(str(cell.value)) > max_length:
                max_length = len(str(cell.value))
        except:
            pass
    adjusted_width = (max_length + 2) * 2
    worksheet.column_dimensions[column_name].width = adjusted_width

# Start the loop to read URLs from somewhere
now = datetime.datetime.now()
df = pd.read_excel("C:\\AlphaData\\NextGen\\NextGen TechLeads.xlsx", sheet_name='04-03-2023')
df['Time Stamp'] = pd.to_datetime(df['Time Stamp']).dt.strftime('%m/%d/%Y')
df['Date of Birth (mm/dd/yy)'] = pd.to_datetime(df['Date of Birth (mm/dd/yy)']).dt.strftime('%m/%d/%Y')

# Use the concurrent.futures library to process the URL links in parallel
with concurrent.futures.ThreadPoolExecutor() as executor:
    futures = []
    for index, row in df.iterrows():
        link = row[11]
        mbinumber = row[6]

        # Submit the URL link to
        futures.append(executor.submit(download_recording, link, index))

    future = concurrent.futures.as_completed(futures)
    for index, result in enumerate(future):
        filename = result.result()
        if filename:
            worksheet.append([df.iloc[index][0], df.iloc[index][1], df.iloc[index][2], df.iloc[index][3], df.iloc[index][4], df.iloc[index][5], df.iloc[index][6], df.iloc[index][7], df.iloc[index][8], df.iloc[index][9],df.iloc[index][10], "N/A", df.iloc[index][11], now.strftime("%Y-%m-%d %I:%M:%S %p"),"YES", ""])
        else:
            worksheet.append([df.iloc[index][0], df.iloc[index][1], df.iloc[index][2], df.iloc[index][3], df.iloc[index][4], df.iloc[index][5], df.iloc[index][6], df.iloc[index][7], df.iloc[index][8], df.iloc[index][9],df.iloc[index][10], "N/A", df.iloc[index][11], now.strftime("%Y-%m-%d %I:%M:%S %p"),"NO", "Error Downloading caused by Connection or Timeout Error!"])

        # Save the workbook after every download
        workbook.save('NG_04-03-2023_ALPHA.xlsx')

shutil.move("C:\\AlphaData\\Recordings\\NextGen\\NG_04-03-2023_ALPHA.xlsx","C:\\AlphaData\\Recordings\\NextGen\\April\\3\\")

workbook.close()
