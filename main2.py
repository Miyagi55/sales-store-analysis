import pandas as pd
from concurrent.futures import ThreadPoolExecutor
from functools import partial
import os

def read_sheet(file_path, sheet_name, skiprows=42, usecols='G'):
    """Helper function to read a single sheet."""
    try:
        return sheet_name, pd.read_excel(
            file_path,
            sheet_name=sheet_name,
            skiprows=skiprows,
            usecols=usecols,
            header=None,
            engine='openpyxl'
        )
    except ValueError as e:
        print(f"Error reading sheet '{sheet_name}': {e}")
        return sheet_name, None

def process_excel_file(file_path, sheets_to_read):
    """Process single Excel file with multiple sheets."""
    read_sheet_partial = partial(read_sheet, file_path)
    
    with ThreadPoolExecutor(max_workers=min(len(sheets_to_read), 4)) as executor:
        results = executor.map(read_sheet_partial, sheets_to_read)
    
    return {name: df for name, df in results if df is not None}

def calculate_monthly_averages(files_dir, months, sheets_to_read):
    """Calculate monthly averages for all stores."""
    monthly_averages = {}
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    files_path = os.path.join(script_dir, files_dir)
    
    for month in months:
        file_name = f'ECUADOR - {month}_ 2024.xlsx'
        file_path = os.path.join(files_path, file_name)
        
        try:
            print(f"Processing file: {file_path}")
            data_frames = process_excel_file(file_path, sheets_to_read)
            
            store_totals = {}
            for store, df in data_frames.items():
                total = int(round(df[df.columns[0]].sum()))
                store_totals[store] = total
            
            monthly_avg = int(round(sum(store_totals.values()) / len(store_totals)))
            monthly_averages[month] = {
                'average': monthly_avg,
                'store_totals': store_totals
            }
            
        except Exception as e:
            print(f"Error processing {file_name}: {e}")
            continue
    
    return monthly_averages

def create_summary_excel(monthly_averages, output_path):
    """Create summary Excel file with monthly averages and store details."""
    # Monthly averages sheet remains the same
    avg_data = {month: data['average'] for month, data in monthly_averages.items()}
    avg_df = pd.DataFrame([avg_data], index=['Monthly Average'])
    
    # Reorganize store details with months as rows and stores as columns
    store_data = []
    # Get list of all stores (columns)
    stores = list(next(iter(monthly_averages.values()))['store_totals'].keys())
    
    # Create data for each month
    for month in monthly_averages.keys():
        row_data = {'Month': month}  # Add month column
        for store in stores:
            row_data[store] = monthly_averages[month]['store_totals'].get(store, 0)
        store_data.append(row_data)
    
    # Create DataFrame with Month as index
    store_df = pd.DataFrame(store_data)
    store_df.set_index('Month', inplace=True)
    
    script_dir = os.path.dirname(os.path.abspath(__file__))
    full_output_path = os.path.join(script_dir, output_path)
    
    with pd.ExcelWriter(full_output_path, engine='openpyxl') as writer:
        avg_df.to_excel(writer, sheet_name='Monthly Averages')
        store_df.to_excel(writer, sheet_name='Store Details')
        
        # Auto-adjust column widths
        for sheet_name in writer.sheets:
            worksheet = writer.sheets[sheet_name]
            for column in worksheet.columns:
                max_length = 0
                column = [cell for cell in column]
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(cell.value)
                    except:
                        pass
                adjusted_width = (max_length + 2)
                worksheet.column_dimensions[column[0].column_letter].width = adjusted_width

if __name__ == '__main__':
    files_dir = 'files'
    months = ['FEBRERO', 'MARZO', 'ABRIL', 'MAYO', 'JUNIO', 'JULIO', 
              'AGOSTO', 'SEPTIEMBRE', 'OCTUBRE', 'NOVIEMBRE']
    sheets_to_read = ["CUENCA", "MALL DEL SOL", "MALL DEL NORTE", "DORADO",
                     "ENTRERIOS", "CEIBOS", "M PACIFICO", "PORTOVIEJO",
                     "PASEO MANTA", "LA LIBERTAD", "BABAHOYO"]
    output_path = 'monthly_sales_summary_2024.xlsx'
    
    monthly_averages = calculate_monthly_averages(files_dir, months, sheets_to_read)
    create_summary_excel(monthly_averages, output_path)
    
    print(f"Summary has been saved to {output_path}")