import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates #for handing date/time formatting on charts
import seaborn as sns
from datetime import datetime, timedelta

def analyze_air_quality_data(file_path):
    """
    Analyzes air purifier data from an Excel file.

    This function performs the following steps:
    1. Loads the data from the specified Excel file.
    2. Calculates and prints a detailed statistical summary.
    3. Visualizes the data with a time-series plot and histograms.
    """

    # Load the Excel file
    df = pd.read_excel(file_path) #converts to a dataframe

    # Columns from the excel sheet (same for consistency)
    df.columns = ['Oxygen_Concentration','AQI']
    print(f"Successfully loaded data from '{file_path}'.")

    start_time = datetime.now().replace(hour=0, minute=0, second=0, microsecond=0)
    time_intervals = len(df)
    
    if time_intervals <= 1440:
        # If we have 1440 or fewer points, spread them across 24 hours
        minutes_per_interval = max(1, 1440 // time_intervals)
        df.index = pd.date_range(start=start_time, periods=time_intervals, freq=f'{minutes_per_interval}min') 
    else:
        # If we have more than 1440 points, use seconds
        seconds_per_interval = max(1, 86400 // time_intervals)  # 86400 seconds in a day
        df.index = pd.date_range(start=start_time, periods=time_intervals, freq=f'{seconds_per_interval}s')
    
    # --- Statistical Analysis ---
    print("\n--- Step 1: Statistical Summary ---")
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
    print(f"- Mean O2: {statistical_summary.loc['mean', 'Oxygen_Concentration']:.2f}%. Normal range is 19-21%.")

    # --- Visualization ---
    df_resampled = df.resample('T').mean()
    print("\nNote: Resampling data to 1-minute averages for cleaner time-series plots.")

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
    ax2 = sns.lineplot(x=df_resampled.index, y=df_resampled['Oxygen_Concentration'], color='tab:red', linewidth=2.5)
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
    sns.histplot(df['Oxygen_Concentration'], kde=True, color='skyblue')
    plt.title('Oxygen Concentration Distribution', fontsize=14)
    plt.xlabel('Oxygen Concentration (%)')
    plt.ylabel('Frequency')
    
    #creating an image and saving it
    plt.tight_layout(rect=[0, 0, 1, 0.96])
    plt.savefig('distributions_plot.png', dpi=300, bbox_inches='tight')
    print("Saved the distributions plot as 'distributions_plot.png'")
    
    # 3. Correlation Analysis
    print("\n--- Step 3: Correlation Analysis ---")
    correlation = df['AQI'].corr(df['Oxygen_Concentration'])
    print(f"Correlation between AQI and Oxygen Concentration: {correlation:.3f}")
    if abs(correlation) > 0.7:
        print("Strong correlation detected!")
    elif abs(correlation) > 0.3:
        print("Moderate correlation detected.")
    else:
        print("Weak correlation - these variables are mostly independent.")
    
    print("\n--- Analysis Complete ---")
    print("Check the console for statistics and look for the saved .png files for your plots.")
    
    return df  # Return the dataframe for further analysis if needed


if __name__ == '__main__':
    analyze_air_quality_data('Oxygen_AQI_Dataset.xlsx')