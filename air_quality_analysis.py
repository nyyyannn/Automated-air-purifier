import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import matplotlib.dates as mdates
import seaborn as sns

def analyze_air_quality_data(file_path):
    """
    Analyzes air purifier data from an Excel file.

    This function performs the following steps:
    1. Loads the data from the specified Excel file.
    2. Calculates and prints a detailed statistical summary.
    3. Visualizes the data with a time-series plot and histograms.
    """
    print("--- Starting Air Quality Data Analysis ---")

    # Load the Excel file
    df = pd.read_excel(file_path)

    # Standardize column names for consistency
    df.columns = ['AQI', 'Oxygen_Concentration']
    print(f"Successfully loaded data from '{file_path}'.")

    # --- Statistical Analysis ---
    print("\n--- Step 1: Statistical Summary ---")
    statistical_summary = df.describe()
    print("The table below shows the key statistical metrics for your data:")
    print(statistical_summary)
    print("\nInterpretation:")
    print(f"- Mean AQI: {statistical_summary.loc['mean', 'AQI']:.2f}. This is the average pollution level.")
    print(f"- Std Dev AQI: {statistical_summary.loc['std', 'AQI']:.2f}. This shows how much the AQI fluctuated. A high value means many spikes.")
    print(f"- Max AQI: {statistical_summary.loc['max', 'AQI']:.2f}. This was the worst air quality moment.")
    print(f"- Mean O2: {statistical_summary.loc['mean', 'Oxygen_Concentration']:.2f}%. Normal range is 19-21%.")

    # --- Visualization ---
    print("\n--- Step 2: Generating Visualizations ---")
    
    sns.set_theme(style="whitegrid")

    # 1. Time-Series Plot
    fig, ax1 = plt.subplots(figsize=(16, 8))
    fig.suptitle('Air Purifier Sensor Readings Over Time', fontsize=18, weight='bold')

    # Plot AQI on the primary y-axis
    color = 'tab:red'
    ax1.set_xlabel('Time', fontsize=12)
    ax1.set_ylabel('Air Quality Index (AQI)', color=color, fontsize=14)
    ax1.plot(df.index, df['AQI'], color=color, label='AQI', linewidth=2)
    ax1.tick_params(axis='y', labelcolor=color)

    # Create a second y-axis for O2 Concentration
    ax2 = ax1.twinx()
    color = 'tab:blue'
    ax2.set_ylabel('Oxygen Concentration (%)', color=color, fontsize=14)
    ax2.plot(df.index, df['Oxygen_Concentration'], color=color, label='Oxygen (%)', linewidth=2)
    ax2.tick_params(axis='y', labelcolor=color)
    
    # Format the time on the x-axis to be readable (if datetime index exists)
    try:
        ax1.xaxis.set_major_formatter(mdates.DateFormatter('%H:%M'))
        plt.setp(ax1.xaxis.get_majorticklabels(), rotation=45)
    except:
        pass  # Use default formatting if not datetime index
    
    plt.tight_layout(rect=[0, 0, 1, 0.96])
    plt.savefig('time_series_plot.png', dpi=300, bbox_inches='tight')
    print("Saved the time-series plot as 'time_series_plot.png'")

    # 2. Histograms
    plt.figure(figsize=(16, 6))
    plt.suptitle('Distribution of Sensor Values', fontsize=18, weight='bold')
    
    plt.subplot(1, 2, 1)
    sns.histplot(df['AQI'], kde=True, color='salmon')
    plt.title('AQI Distribution', fontsize=14)
    plt.xlabel('Air Quality Index')
    plt.ylabel('Frequency')
    
    plt.subplot(1, 2, 2)
    sns.histplot(df['Oxygen_Concentration'], kde=True, color='skyblue')
    plt.title('Oxygen Concentration Distribution', fontsize=14)
    plt.xlabel('Oxygen Concentration (%)')
    plt.ylabel('Frequency')
    
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