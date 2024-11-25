import pandas as pd
from datetime import datetime, timedelta
import os
import statistics

project_root = os.path.dirname(os.path.abspath(__file__))
output_path = os.path.join(project_root, 'summary_results.xlsx')

def analyze_prices(directory):
    # Initialize lists to store results
    dates = []
    avg_prices = []

    # Get all Excel files in the directory
    excel_files = [f for f in os.listdir(directory) if f.endswith('.xlsx')]

    # Process each Excel file
    for excel_file in excel_files:

        # Parse the Excel file
        df = pd.ExcelFile(os.path.join(directory, excel_file))

        # Extract data for hours 2-15
        for hour in range(2, 16):
            sheet_name = f'{hour}'

            try:
                hourly_df = df.parse(sheet_name)

                # Filter for the Republic of Buryatia
                buriatia_index = hourly_df.iloc[:, 4] == 'Республика Бурятия'
                if not buriatia_index.any():
                    continue

                # Find the column index containing price data
                price_column = hourly_df.columns[-2]

                buriatia_prices = hourly_df.loc[buriatia_index, price_column]

                # Convert prices to float and calculate average
                price_series = pd.to_numeric(buriatia_prices, errors='coerce')
                valid_prices = price_series.dropna()

                if len(valid_prices) == 0:
                    print(f"No valid prices found for Buryatia in sheet {sheet_name}")
                else:
                    avg_price = valid_prices.mean()

                    dates.append(extract_date_from_filename(excel_file))
                    avg_prices.append(avg_price)

                    print(f"Processed sheet {sheet_name}: Average price = {avg_price:.2f} rub/MWh")
            except Exception as e:
                print(f"Error processing sheet {sheet_name} in file {excel_file}: {str(e)}")

    # Create summary DataFrame
    summary_df = pd.DataFrame({'Date': dates, 'Average Price': avg_prices})

    # Save results to Excel file
    output_filename = 'summary_results.xlsx'
    summary_df.to_excel(output_path, index=False)

    print(f"\nРезультаты сохранены в {output_path}")


def extract_date_from_filename(filename):
    parts = filename.split('_')
    if len(parts) >= 2:
        return parts[1].split('.')[0]
    return None


if __name__ == "__main__":
    download_directory = os.path.join(project_root, 'downloaded_xls')
    analyze_prices(download_directory)

