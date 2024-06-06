import pandas as pd
import heapq

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

def calculate_priority(metrics):
    weights = {'import_ton': 1, 'export_ton': 2, 'import_usd': 3, 'export_usd': 4}
    priority = (weights['import_ton'] * metrics['import_ton'] +
                weights['export_ton'] * metrics['export_ton'] +
                weights['import_usd'] * metrics['import_usd'] +
                weights['export_usd'] * metrics['export_usd'])
    return priority

def top_5_by_priority(country_data):
    continent_groups = {}
    for country, metrics in country_data.items():
        continent = metrics['continent']
        if continent not in continent_groups:
            continent_groups[continent] = []
        priority = calculate_priority(metrics)
        heapq.heappush(continent_groups[continent], (-priority, country, metrics))

    top_5_countries = {}
    for continent, countries in continent_groups.items():
        top_5_countries[continent] = [heapq.heappop(countries) for _ in range(min(5, len(countries)))]

    return top_5_countries

def main():
    filename = 'dataFIXmerges.xlsx'
    country_data = read_data(filename)
    if not country_data:
        print("No valid data found.")
        return
    
    top_5_countries = top_5_by_priority(country_data)

    for continent, countries in top_5_countries.items():
        print(f"Top 5 Countries in {continent} by Priority Score:")
        for priority, country, metrics in countries:
            print(f"{country}: Priority Score = {-priority}")
        print("\n")

if __name__ == "__main__":
    main()
