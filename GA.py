import pandas as pd
import numpy as np
import random
from deap import base, creator, tools, algorithms

# Load the Excel file
file_path = 'dataFIXmerges.xlsx'
excel_data = pd.ExcelFile(file_path)

# Load data from Sheet1
sheet1_data = pd.read_excel(file_path, sheet_name='Sheet1')

# Clean and preprocess the data
sheet1_data = sheet1_data.drop(0)  # Remove the first row which seems to be a header for the columns
sheet1_data.columns = ['Country', 'Continent', 'Export_2018_Vol', 'Export_2019_Vol', 'Export_2020_Vol', 'Export_2021_Vol', 'Export_2022_Vol', 'Export_2023_Vol',
                       'Extra1', 'Extra2', 'Extra3', 'Extra4', 'Extra5', 'Extra6', 'Extra7', 'Extra8', 'Extra9', 'Extra10', 'IMPOR', 'Unnamed: 19', 'Unnamed: 20', 
                       'Unnamed: 21', 'Unnamed: 22', 'Unnamed: 23', 'Unnamed: 24', 'Unnamed: 25', 'Unnamed: 26', 'Unnamed: 27', 'Unnamed: 28', 'Unnamed: 29', 
                       'Unnamed: 30', 'Unnamed: 31', 'Unnamed: 32', 'Unnamed: 33', 'Unnamed: 34']

# Drop unnecessary columns with 'Extra' in their names and other unnamed columns
columns_to_drop = [col for col in sheet1_data if 'Extra' in col or 'Unnamed' in col]
sheet1_data = sheet1_data.drop(columns=columns_to_drop)

# Convert numerical columns to proper data types
for col in sheet1_data.columns[2:]:
    sheet1_data[col] = pd.to_numeric(sheet1_data[col], errors='coerce')

# Define the genetic algorithm components
creator.create("FitnessMax", base.Fitness, weights=(1.0,))
creator.create("Individual", list, fitness=creator.FitnessMax)

def create_individual():
    return creator.Individual(np.random.permutation(len(sheet1_data)).tolist())

def calculate_fitness(individual):
    total_export = sheet1_data.iloc[individual, 2:8].sum().sum()  # Export columns from 2018 to 2023
    total_import = sheet1_data.iloc[individual, -6:].sum().sum()  # Import columns from 2018 to 2023
    return total_export + total_import,

toolbox = base.Toolbox()
toolbox.register("individual", create_individual)
toolbox.register("population", tools.initRepeat, list, toolbox.individual)
toolbox.register("evaluate", calculate_fitness)
toolbox.register("mate", tools.cxTwoPoint)
toolbox.register("mutate", tools.mutShuffleIndexes, indpb=0.05)
toolbox.register("select", tools.selTournament, tournsize=3)

# # Genetic Algorithm parameters
population_size = 50
generations = 200
mutation_rate = 0.2

# Initialize the population
population = toolbox.population(n=population_size)

# Run the Genetic Algorithm
algorithms.eaSimple(population, toolbox, cxpb=0.7, mutpb=mutation_rate, ngen=generations, verbose=True)

# Get the best solution
best_individual = tools.selBest(population, k=1)[0]
best_fitness = calculate_fitness(best_individual)

# Get the data corresponding to the best individual
best_data = sheet1_data.iloc[best_individual].reset_index(drop=True)

# Add the total export and import as a new column
best_data['Total_Export_Import'] = best_data.iloc[:, 2:].sum(axis=1)

# Get top 5 countries by total export and import for each continent
top5_per_continent = best_data.groupby('Continent').apply(lambda x: x.nlargest(5, 'Total_Export_Import')).reset_index(drop=True)

# Save the top 5 per continent to a new Excel file
output_file_path = 'top5_exports_imports_by_continent.xlsx'
top5_per_continent.to_excel(output_file_path, index=False)

# Print the best solution and its fitness
print("Best Individual:", best_individual)
print("Best Fitness:", best_fitness[0])
print(f"Top 5 countries by export and import volume per continent saved to {output_file_path}")

# Display the cleaned data (optional)
print(sheet1_data)
