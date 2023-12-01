import os
from bs4 import BeautifulSoup
import pandas as pd
import gspread
from google.oauth2.service_account import Credentials

def extract_data_from_html(file_path, call_type):
    with open(file_path, 'r', encoding='utf-8') as file:
        content = file.read()
        soup = BeautifulSoup(content, 'html.parser')

        abbr_element = soup.find('abbr', class_='published')
        date_time = abbr_element.get_text(strip=True).replace('\u202f', ' ') if abbr_element else None
        
        if date_time and '\nCentral Time' in date_time:
            date_time = date_time.replace('\nCentral Time', '')

        if call_type == 'Received':
            duration_element = soup.find('abbr', class_='duration')
            duration = duration_element.get_text(strip=True) if duration_element else None
        else:
            duration = None
        
        return date_time, duration

def classify_call_type(filename):
    if ("voicemail" in filename.lower() or "missed" in filename.lower()) and '+16502651193' not in filename:
        return "Missed call"
    elif "received" in filename.lower() and '+16502651193' not in filename:
        return "Received"

def get_user_input():
    while True:
        try:
            worksheet_title = input("Enter the client code: ")
            start_date = pd.to_datetime(input("Enter the first day of the reporting period (YYYY-M-D): "))
            end_date = pd.to_datetime(input("Enter the last day of the reporting period (YYYY-M-D): "))
            break
        except ValueError:
            print("Invalid input. Please enter a valid name and date format (YYYY-MM-DD).")
    return worksheet_title, start_date, end_date

def calculate_additional_columns(filtered_df):
    total_calls = len(filtered_df)
    missed_calls = len(filtered_df[filtered_df['Call Type'] == 'Missed call'])
    pickups = len(filtered_df[filtered_df['Call Type'] == 'Received'])
    pickup_rate = pickups / total_calls * 100 if total_calls > 0 else 0

    return total_calls, missed_calls, pickups, pickup_rate

# Update the clear_specific_columns function
def clear_specific_columns(worksheet):
    cols_to_clear = ['D', 'E', 'F', 'G']  # Columns to clear from D to G

    # Get the entire range of data in columns D through G starting from row 3
    cells_to_clear = worksheet.range('D3:G' + str(worksheet.row_count))

    # Clear the contents of cells in the specified range
    for cell in cells_to_clear:
        cell.value = ''

    # Update the cells in the worksheet
    worksheet.update_cells(cells_to_clear)

def remove_empty_rows(worksheet):
    # Fetch all values from the worksheet
    all_values = worksheet.get_all_values()

    # Find rows with no values in any cell and create a list of indices to delete
    rows_to_delete = [index + 1 for index, row in enumerate(all_values) if not any(row)]

    # Delete the identified rows
    for row_index in reversed(rows_to_delete):  # Deleting in reverse to avoid index shifting
        worksheet.delete_row(row_index)


def process_folder(folder_path):
    data = []
    for file_name in os.listdir(folder_path):
        if file_name.endswith(".html"):
            file_path = os.path.join(folder_path, file_name)
            call_type = classify_call_type(file_name)
            date_time, duration = extract_data_from_html(file_path, call_type)
            data.append({
                'Call Type': call_type,
                'Date and Time': date_time,
                'Duration': duration
            })

    df = pd.DataFrame(data)
    df['Date and Time'] = pd.to_datetime(df['Date and Time'])
    df.sort_values(by='Date and Time', inplace=True)
    df.dropna(subset=['Call Type', 'Date and Time'], inplace=True)

    print("DataFrame created. Proceeding to Google Sheets...")

    worksheet_title, start_date, end_date = get_user_input()

    filtered_df = df[
        (df['Date and Time'] >= start_date) &
        (df['Date and Time'] <= end_date) &
        ((df['Call Type'] == 'Received') | (df['Call Type'] == 'Missed call'))
    ]

    try:
        worksheet_title_base = f"{start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}"
        worksheet_title = worksheet_title + " - Call Report " + worksheet_title_base
        counter = 1

        credentials = Credentials.from_service_account_file(
          #Enter the file path of the json file for the google sheets api
            '',
            scopes=[
                "https://www.googleapis.com/auth/spreadsheets",
                "https://www.googleapis.com/auth/drive"
            ]
        )
        client = gspread.authorize(credentials)
        
        new_spreadsheet = client.create(worksheet_title)
        worksheet = new_spreadsheet.get_worksheet(0)

        total_calls, missed_calls, pickups, pickup_rate = calculate_additional_columns(filtered_df)

        filtered_df['Total Calls'] = total_calls
        filtered_df['Missed Calls'] = missed_calls
        filtered_df['Pickups'] = pickups
        filtered_df['Pickup Rate'] = f"{pickup_rate:.2f}%"

        clear_specific_columns(worksheet)

        filtered_df = filtered_df.astype(str)
        worksheet.update([filtered_df.columns.values.tolist()] + filtered_df.values.tolist())

        remove_empty_rows(worksheet)

        print(f"Data within the specified date range saved to '{worksheet_title}' on Google Sheets.")

      # Entere the email you want to share the file with
        new_spreadsheet.share("", perm_type='user', role='writer')
        print(f"Spreadsheet URL: {new_spreadsheet.url}")

    except Exception as e:
        print("Error occurred while updating Google Sheets:", e)
  
# Enter the file path where the data belongs
folder_path = '/Users/Name/Desktop/Takeout/Voice/Calls'
process_folder(folder_path)
