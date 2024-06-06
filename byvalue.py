import pandas as pd

def read_data(filename):
    country_data = {}
    try:
        df = pd.read_excel(filename, skiprows=1)
    except Exception as e:
        print(f"Error reading Excel file: {e}")
        return country_data

    for index, row in df.iterrows():
        if pd.notna(row.iloc[0]) and isinstance(row.iloc[0], str):
            country = row.iloc[0].strip()
            continent = row.iloc[1].strip() if pd.notna(row.iloc[1]) else "Unknown"
            metrics = {}
            try:
                import_ton = pd.to_numeric(row.iloc[2], errors='coerce') 
                export_ton = pd.to_numeric(row.iloc[3], errors='coerce') 
                import_usd = pd.to_numeric(row.iloc[33], errors='coerce')
                export_usd = pd.to_numeric(row.iloc[10], errors='coerce')

                metrics['import_ton'] = import_ton if pd.notna(import_ton) else 0
                metrics['export_ton'] = export_ton if pd.notna(export_ton) else 0
                metrics['import_usd'] = import_usd if pd.notna(import_usd) else 0
                metrics['export_usd'] = export_usd if pd.notna(export_usd) else 0

                metrics['continent'] = continent
                
                if any([metrics['import_ton'], metrics['export_ton'], metrics['import_usd'], metrics['export_usd']]):
                    country_data[country] = metrics
            except ValueError as ve:
                print(f"Invalid value in row {index + 2}: {ve}")
                continue
    return country_data

def top_and_bottom_5_countries_by_value_or_ton(country_data, metric):
    continent_groups = {}
    for country, metrics in country_data.items():
        continent = metrics['continent']
        if continent not in continent_groups:
            continent_groups[continent] = []
        continent_groups[continent].append((country, metrics))

    top_and_bottom_5_countries = {}
    for continent, countries in continent_groups.items():
        top_countries = sorted(countries, key=lambda x: x[1][metric], reverse=True)[:5]
        bottom_countries = sorted(countries, key=lambda x: x[1][metric])[:5]
        top_and_bottom_5_countries[continent] = {
            'top_countries': top_countries,
            'bottom_countries': bottom_countries
        }
    
    return top_and_bottom_5_countries

def main():
    filename = 'dataFIXmerges.xlsx'
    country_data = read_data(filename)
    if not country_data:
        print("No valid data found.")
        return
    
    metrics = ['import_usd', 'export_usd', 'import_ton', 'export_ton']
    for metric in metrics:
        top_and_bottom_5_countries = top_and_bottom_5_countries_by_value_or_ton(country_data, metric)
        print(f"Top and Bottom 5 countries by {metric}:")
        for continent, countries in top_and_bottom_5_countries.items():
            print(f"Continent: {continent}")
            
            print(f"Top 5 {metric} Countries:")
            for country, metrics in countries['top_countries']:
                print(f"{country}: {metric} = {metrics[metric]}")
            print("\n")
            
            print(f"Bottom 5 {metric} Countries:")
            for country, metrics in countries['bottom_countries']:
                print(f"{country}: {metric} = {metrics[metric]}")
            print("\n")

if __name__ == "__main__":
    main()
