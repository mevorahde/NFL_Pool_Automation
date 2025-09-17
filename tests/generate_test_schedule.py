import os
import pandas as pd

# Ensure the 'tests' folder exists
os.makedirs("test_runs", exist_ok=True)

# Define test data for edge-case NFL games
data = [
    {"Excel_Row": 1, "Team1": "DET", "Team2": "CHI", "Expected_Weekday": "Thursday"},
    {"Excel_Row": 2, "Team1": "DAL", "Team2": "NYG", "Expected_Weekday": "Thursday"},
    {"Excel_Row": 3, "Team1": "SF",  "Team2": "SEA", "Expected_Weekday": "Thursday"},
    {"Excel_Row": None, "Team1": "NYJ", "Team2": "MIA", "Expected_Weekday": "Friday"},
    {"Excel_Row": 4, "Team1": "BUF", "Team2": "NE",  "Expected_Weekday": "Saturday"},
    {"Excel_Row": 5, "Team1": "CIN", "Team2": "BAL", "Expected_Weekday": "Saturday"},
    {"Excel_Row": 6, "Team1": "LAC", "Team2": "LV",  "Expected_Weekday": "Saturday"},
]

# Create DataFrame
df = pd.DataFrame(data)

# Show contents for confirmation
print("\n📋 DataFrame preview:")
print(df)

# Save to Excel
output_path = os.path.abspath("test_runs/test_schedule.xlsx")
df.to_excel(output_path, index=False)

# Confirm success
print(f"\n✅ Excel file saved to: {output_path}")