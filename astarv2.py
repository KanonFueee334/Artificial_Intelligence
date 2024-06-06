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
                import_usd = pd.to_numeric(row.iloc[27], errors='coerce')
                export_usd = pd.to_numeric(row.iloc[8], errors='coerce')

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

def calculate_priority(country_data):
    weights = {'import_ton': 1, 'export_ton': 2, 'import_usd': 3, 'export_usd': 4}
    for country, metrics in country_data.items():
        metrics['priority'] = (weights['import_ton'] * metrics['import_ton'] +
                               weights['export_ton'] * metrics['export_ton'] +
                               weights['import_usd'] * metrics['import_usd'] +
                               weights['export_usd'] * metrics['export_usd'])
    return country_data

def top_and_bottom_countries_by_continent(country_data, n=5):
    continent_groups = {}
    for country, metrics in country_data.items():
        continent = metrics['continent']
        if continent not in continent_groups:
            continent_groups[continent] = []
        continent_groups[continent].append((country, metrics))

    top_and_bottom_countries = {}
    for continent, countries in continent_groups.items():
        top_import_ton = sorted(countries, key=lambda x: x[1]['import_ton'], reverse=True)[:n]
        bottom_import_ton = sorted(countries, key=lambda x: x[1]['import_ton'])[:n]
        top_export_ton = sorted(countries, key=lambda x: x[1]['export_ton'], reverse=True)[:n]
        bottom_export_ton = sorted(countries, key=lambda x: x[1]['export_ton'])[:n]
        top_priority = sorted(countries, key=lambda x: x[1]['priority'], reverse=True)[:n]
        bottom_priority = sorted(countries, key=lambda x: x[1]['priority'])[:n]
        top_and_bottom_countries[continent] = {
            'top_import_ton': top_import_ton,
            'bottom_import_ton': bottom_import_ton,
            'top_export_ton': top_export_ton,
            'bottom_export_ton': bottom_export_ton,
            'top_priority': top_priority,
            'bottom_priority': bottom_priority
        }
    
    return top_and_bottom_countries

def main():
    filename = 'dataFIXmerges.xlsx'
    country_data = read_data(filename)
    if not country_data:
        print("No valid data found.")
        return
    
    country_data = calculate_priority(country_data)
    top_and_bottom_countries = top_and_bottom_countries_by_continent(country_data)

    for continent, countries in top_and_bottom_countries.items():
        print(f"Top {len(countries['top_import_ton'])} Countries in {continent} by Import Ton:")
        for country, metrics in countries['top_import_ton']:
            print(f"{country}: Import Ton = {metrics['import_ton']}")
        print("\n")
        
        print(f"Bottom {len(countries['bottom_import_ton'])} Countries in {continent} by Import Ton:")
        for country, metrics in countries['bottom_import_ton']:
            print(f"{country}: Import Ton = {metrics['import_ton']}")
        print("\n")
        
        print(f"Top {len(countries['top_export_ton'])} Countries in {continent} by Export Ton:")
        for country, metrics in countries['top_export_ton']:
            print(f"{country}: Export Ton = {metrics['export_ton']}")
        print("\n")
        
        print(f"Bottom {len(countries['bottom_export_ton'])} Countries in {continent} by Export Ton:")
        for country, metrics in countries['bottom_export_ton']:
            print(f"{country}: Export Ton = {metrics['export_ton']}")
        print("\n")
        
        print(f"Top {len(countries['top_priority'])} Countries in {continent} by Priority Score:")
        for country, metrics in countries['top_priority']:
            print(f"{country}: Priority Score = {metrics['priority']}")
        print("\n")
        
        print(f"Bottom {len(countries['bottom_priority'])} Countries in {continent} by Priority Score:")
        for country, metrics in countries['bottom_priority']:
            print(f"{country}: Priority Score = {metrics['priority']}")
        print("\n")

if __name__ == "__main__":
    main()
