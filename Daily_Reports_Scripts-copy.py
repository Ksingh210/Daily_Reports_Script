import pandas as pd

def main():

    # Insert Date that match report file names
    report_date = str(input("Today's Date (Month.Day): "))

    # The remove week is to remove past fiscal weeks from the forecast sum below. Must increase by 1 per week.
    remove_wk = 10

    # File paths set for location of files
    amt_file_path = (r'#')
    inventory_file_path = (r'#')
    sales_forecast_path = (r'#')
    kgp_pos_units_path = (r'#')

    # Variable assignment and import for files
    amt = pd.read_excel(amt_file_path)
    inventory = pd.read_excel(inventory_file_path)
    sales_forecast = pd.read_excel(sales_forecast_path)
    units = pd.read_excel(kgp_pos_units_path)

    # Selecting Fiscal Weeks for sum total
    sales_forecast_weeks = list(sales_forecast)
    del sales_forecast_weeks[0:remove_wk]

    # New column for sum total of fiscal week selections 
    sales_forecast['THD Forecast'] = sales_forecast[sales_forecast_weeks].sum(axis=1)
    sales_forecast.rename(columns={'STR_NBR': 'STR NBR',
                                'SKU_NBR': 'Sku'}, inplace=True)

    # New sales_forecast table with only necessary columns
    columnslist = ['STR NBR', 'Sku', 'Forecast']
    sales_forecast = sales_forecast[columnslist]

    # Name changes to match amt file column names
    inventory.rename(columns={'Store Nbr': 'STR NBR',
                            'SKU Nbr': 'Sku',
                            'Str OH Units Dly': 'OH'}, inplace = True)

    # Removing unnecessary columns
    inventory.drop(columns = ['SKU Name', 'Str OO Units Dly'], inplace = True) 

    # Name change to match amt file column names
    units.rename(columns={'Store': 'STR NBR',
                                'Item': 'Sku',
                                2018: 'Units'}, inplace=True)

    # Merging OH data to AMT
    combo = amt.merge(inventory,
                    how='left',
                    on=['STR NBR', 'Sku'])

    # Second merge, deleting unnecessary columns, and renaming to match AMT
    combo_b = combo.merge(units,
                        how='left',
                        left_on=['STR NBR', 'Old Sku'],
                        right_on=['STR NBR', 'Sku'])
    del combo_b['Sku_y']
    combo_b.rename(columns={'Sku_x': 'Sku'}, inplace=True)

    # Merging forecast to new OH and AMT
    daily_reports = combo_b.merge(sales_forecast,
                                how='left',
                                on=['STR NBR', 'Sku'])
    daily_reports = daily_reports.fillna(0)

    # Insert OH PLTs column in index position 12
    daily_reports.insert(12,'OH PLTs',round(daily_reports['OH'] / daily_reports['PLT SZ']).astype('int'))

    # AWS For KGP Data
    daily_reports['KGP AWS'] = daily_reports['Units'] / len(sales_forecast_weeks)

    # AWS For THD Data
    daily_reports['THD AWS'] = daily_reports['Forecast'] / len(sales_forecast_weeks)

    export = pd.ExcelWriter(r'#\19 '+ report_date +' Daily Reports.xlsx',
                        engine = 'xlsxwriter')

    daily_reports.to_excel(export, index=False)

    export.save()

    print("Reports Executed! Check Daily Reports Folder.")

if __name__ == '__main__':
    main()
