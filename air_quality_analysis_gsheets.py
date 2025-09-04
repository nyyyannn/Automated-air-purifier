from datetime import datetime, timedelta
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import matplotlib.dates as mdates
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import sys

def export_stats_to_google_sheet(dataframe, sheet_name):
    """
    Exports a Pandas DataFrame to a specified Google Sheet.

    Args:
        dataframe (pd.DataFrame): The dataframe to export.
        sheet_name (str): The name of the Google Sheet to send the data to.
    """
    print("\n--- Exporting to Google Sheets ---")
    try:
        # Define the scope of access for the APIs. Here it is accessing the list of all spreadsheets and the contents of the files in google drive
        scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
        
        # Authenticate using the downloaded credentials file
        # MAKE SURE 'credentials.json' IS IN THE SAME FOLDER
        creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', scope)
        client = gspread.authorize(creds)

        # Open the spreadsheet
        spreadsheet = client.open(sheet_name)
        
        worksheet_title = 'Statistical Summary'
        try:
            # Try to get the worksheet
            sheet = spreadsheet.worksheet(worksheet_title)
        except gspread.exceptions.WorksheetNotFound:
            # If it doesn't exist, create it
            print(f"Worksheet '{worksheet_title}' not found. Creating it now.")
            sheet = spreadsheet.add_worksheet(title=worksheet_title, rows="100", cols="20")
        
        print(f"Successfully connected to Google Sheet: '{sheet_name}' and worksheet: '{worksheet_title}'")

        # Clear the sheet before writing new data
        sheet.clear()

        # Convert the DataFrame to a list of lists and update the sheet
        # gspread needs the header as well, so we get that first
        header = dataframe.reset_index().columns.values.tolist()
        data_to_upload = dataframe.reset_index().values.tolist()
        
        sheet.update(range_name='A1', values=[header] + data_to_upload)
        
        print("Successfully exported the statistical summary.")
        print(f"Check your Google Sheet: https://docs.google.com/spreadsheets/d/{spreadsheet.id}")

    except FileNotFoundError:
        print("\nERROR: 'credentials.json' not found.") #credentials.json is used to connect to the spreadsheet
        print("Please ensure you have downloaded the JSON key from Google Cloud")
        print("and placed it in the same directory as this script.")
    except gspread.exceptions.SpreadsheetNotFound:
        print(f"\nERROR: Spreadsheet '{sheet_name}' not found.") #if the sheet isn't created yet. Ensure the name is the same as the one on google sheets
        print("Please ensure the sheet exists and you have shared it with the client email:")
        creds_dict = creds.service_account_email
        print(f"--> {creds_dict}")
    except Exception as e:
        print(f"\nAn unexpected error occurred: {e}")
        print("Please double-check your API permissions and sharing settings.")

def analyze_air_quality_data(file_path):
    """
    Analyzes air purifier data from an Excel file.

    This function performs the following steps:
    1. Loads the data from the specified Excel file.
    2. Dynamically creates a datetime index based on the number of records.
    3. Calculates and prints a detailed statistical summary with interpretation.
    4. Visualizes the data with time-series plots and histograms, saving them as image files.
    """
    print("--- Starting Air Quality Data Analysis ---")

    try:
        df = pd.read_excel(file_path)
        #Columns from the excel sheet (same for consistency)
        df.columns = ['O2_Percentage', 'AQI']
        print(f"Successfully loaded data from '{file_path}'.")
    except FileNotFoundError:
        print(f"Warning: File '{file_path}' not found. Generating sample data.")

    start_time = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    time_intervals = len(df)
    
    if time_intervals <= 1440:
        # If we have 1440 or fewer points, assume they are minutes
        minutes_per_interval = max(1, 1440 // time_intervals)
        df.index = pd.date_range(start=start_time, periods=time_intervals, freq=f'{minutes_per_interval}min') 
    else:
        # If we have more than 1440 points, assume they are seconds
        seconds_per_interval = max(1, 86400 // time_intervals)
        df.index = pd.date_range(start=start_time, periods=time_intervals, freq=f'{seconds_per_interval}s')
    
    # --- Step 1: Statistical Analysis ---
    print("\n--- Statistical Summary ---")
    statistical_summary = df.describe()
    """
        describe: includes the count (number of data points), mean, standard deviation,
        minimum and maximum values, 25%, 50%,75% Quartiles (precentiles)
    """
    print("The table below shows the key statistical metrics for your data:")
    print(statistical_summary)
    print("\nInterpretation:")
    #.loc: is used to get the specific summary in the following format: [row,column_name]   
    print(f"- Mean AQI: {statistical_summary.loc['mean', 'AQI']:.2f}. This is the average pollution level.")
    print(f"- Std Dev AQI: {statistical_summary.loc['std', 'AQI']:.2f}. This shows how much the AQI fluctuated. A high value means many spikes.")
    print(f"- Max AQI: {statistical_summary.loc['max', 'AQI']:.2f}. This was the worst air quality moment.")
    print(f"- Mean O2: {statistical_summary.loc['mean', 'O2_Percentage']:.2f}%. Normal range is roughly 20.9%.")

    # --- Visualizations ---
    print("\n--- Generating Visualizations ---")
    df_resampled = df.resample('T').mean()
    print("Note: Resampling data to 1-minute averages for cleaner time-series plots.")

    sns.set_theme(style="whitegrid")

    # 1a. Time-Series Plot for AQI
    plt.figure(figsize=(16, 6))
    ax1 = sns.lineplot(x=df_resampled.index, y=df_resampled['AQI'], color='tab:blue', linewidth=2.5)
    ax1.set_title('Air Quality Index (AQI) Over Time (1-Minute Averages)', fontsize=16, weight='bold')
    ax1.set_xlabel('Time', fontsize=12)
    ax1.set_ylabel('Air Quality Index (AQI)', fontsize=12)
    ax1.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig('aqi_time_series.png', dpi=300)
    print("Saved the AQI time-series plot as 'aqi_time_series.png'")

    # 1b. Time-Series Plot for Oxygen Concentration
    plt.figure(figsize=(16, 6))
    ax2 = sns.lineplot(x=df_resampled.index, y=df_resampled['O2_Percentage'], color='tab:red', linewidth=2.5)
    ax2.set_title('Oxygen Concentration Over Time (1-Minute Averages)', fontsize=16, weight='bold')
    ax2.set_xlabel('Time', fontsize=12)
    ax2.set_ylabel('Oxygen Concentration (%)', fontsize=12)
    ax2.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
    plt.xticks(rotation=45)
    plt.tight_layout()
    plt.savefig('oxygen_time_series.png', dpi=300)
    print("Saved the Oxygen time-series plot as 'oxygen_time_series.png'")

    # 2. Histograms
    plt.figure(figsize=(16, 6))
    plt.suptitle('Distribution of Sensor Values', fontsize=18, weight='bold')
    
    plt.subplot(1, 2, 1) #1x2 grids of plots and this is first one (AQI)
    sns.histplot(df['AQI'], kde=True, color='salmon') 
    plt.title('AQI Distribution', fontsize=14)
    plt.xlabel('Air Quality Index')
    plt.ylabel('Frequency')
    
    plt.subplot(1, 2, 2)
    sns.histplot(df['O2_Percentage'], kde=True, color='skyblue')
    plt.title('Oxygen Concentration Distribution', fontsize=14)
    plt.xlabel('Oxygen Concentration (%)')
    plt.ylabel('Frequency')
    
    #creating an image and saving it
    plt.tight_layout(rect=[0, 0, 1, 0.96])
    plt.savefig('distributions_plot.png', dpi=300, bbox_inches='tight')
    print("Saved the distributions plot as 'distributions_plot.png'")
    
    print("\n--- Analysis Complete ---")
    
    return statistical_summary


if __name__ == '__main__':
    # Define your source file and target sheet name
    source_excel_file = 'Oxygen_AQI_Dataset.xlsx'
    google_sheet_name = 'Automated air purifier analysis results'

    # Run the analysis
    stats_summary = analyze_air_quality_data(source_excel_file)
    
    # Export the results
    if stats_summary is not None:
        export_stats_to_google_sheet(stats_summary, google_sheet_name)

