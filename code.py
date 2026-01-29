import pandas as pd
import psychrolib

# 1. Initialize PsychroLib (Setting unit system to IP or SI)
# Use psychrolib.SI for Celsius/Pascals or psychrolib.IP for Fahrenheit/PSI
psychrolib.SetUnitSystem(psychrolib.SI)

# 2. Load your Excel data
# Assumes Column E is Dry Bulb and Column F is Dew Point
file_path = 'merged_weather-almariyah-2015-2025.xlsx'
df = pd.read_excel(file_path)

# Standard atmospheric pressure at sea level (101325 Pa)
# You can adjust this if your data is from a high-altitude location
# PRESSURE = 101325 

def calculate_wet_bulb(row):
    try:
        # Extract values from Column E (index 4) and F (index 5)
        db = row.iloc[4] 
        dp = row.iloc[5]
        ap = row.iloc[7]
        
        # Calculate Wet Bulb Temperature
        wb = psychrolib.GetTWetBulbFromTDewPoint(db, dp, ap)
        return round(wb, 2)
    except Exception as e:
        return None

# 3. Apply calculation and create the new dataframe
df['Wet Bulb Temperature'] = df.apply(calculate_wet_bulb, axis=1)

# Create a clean output with only the columns you requested
output_columns = [df.columns[4], 'Wet Bulb Temperature', df.columns[5], df.columns[7]]
final_df = df[output_columns]

# 4. Save to new Excel sheet
final_df.to_excel('psychrometric_results.xlsx', index=False)
print("Done! Results saved to psychrometric_results.xlsx")