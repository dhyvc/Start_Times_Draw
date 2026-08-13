import pandas as pd
from bs4 import BeautifulSoup
import requests
import io

def process_html_tables_to_excel():
    # Fetch and parse HTML
    url = "https://cdn.nivut.org.il/uploads/ce3c8aee-e76e-49c1-980c-79400cb44b23_HTML_Start_Times_By_Category.html"
    response = requests.get(url)
    soup = BeautifulSoup(response.content, 'html.parser', from_encoding='utf-8')

    # Initialize Excel writer
    writer = pd.ExcelWriter('start_times_by_category.xlsx', engine='xlsxwriter')

    # Find all h1 headers and subsequent tables
    h1_headers = soup.find_all('h1')
    aggregated_dfs = []

    for header in h1_headers:
        table = header.find_next('table')
        if not table:
            continue
            
        # Extract table title
        class_name = header.get_text(strip=True)
        
        # Read table into DataFrame
        df = pd.read_html(io.StringIO(str(table)), encoding='utf-8')[0]
        
        # Clean column names and drop unnamed columns
        df.columns = [str(col).strip() for col in df.columns]
        df = df.loc[:, ~df.columns.str.contains('unnamed', case=False, na=False)]
        
        # Rename columns
        column_mapping = {
            'SI': 'CARD NUMBER',
            'שם': 'NAME',
            'מספר חבר': 'STNO',
            'שעה': 'START TIME',
            'מועדון': 'CLUB'
        }
        df.rename(columns=column_mapping, inplace=True)
        
        # Add CLASS column
        df['CLASS'] = class_name
        
        # Write to Excel sheet
        df.to_excel(writer, sheet_name=class_name[:31], index=False)  # Sheet name max 31 chars
        aggregated_dfs.append(df)

    # Save aggregated data to CSV with Windows-1255 encoding
    if aggregated_dfs:
        combined_df = pd.concat(aggregated_dfs, ignore_index=True)
        combined_df.to_csv('combined_start_times.csv', encoding='windows-1255', index=False)
    
    writer.close()

if __name__ == "__main__":
    process_html_tables_to_excel()
