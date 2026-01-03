
"""
path: /Users/name/Desktop/
"""

import pandas as pd


week = "Week 1"
sheet_name = "Week 1"

file_path = "/Users/Desktop/Data 2025.xlsx"

data_df = pd.read_excel(file_path, sheet_name="Registration", skiprows=3)
data_df.columns = ["empty", "none1", "none2", "none3", "food", "num", "week", "name", "grade", "sport", "none4", "none5"]

data_filtered = data_df[df["week"] == week]
data = data_filtered[["name", "sport", "food"]]
data = data.reset_index(drop=True)
data = data[data['name'].notna() & (data['name'].astype(str).str.strip() != '')]
data['sport'] = data['sport'].str.strip().str.title()
valid_sports = ['Soccer', 'Dance', 'Badminton', 'Boxing']
data['sport'] = data['sport'].apply(lambda x: x if x in valid_sports else 'N/A')

data['sport'] = data['sport'].replace('', pd.NA)
data['sport'] = data['sport'].fillna('N/A')

data_df = pd.read_excel(file_path, sheet_name=sheet_name, header=None, skiprows=11)
data = df[[2, 4]]
data.columns = ["nameList", "roomNum"]
data = data.dropna(how="all").reset_index(drop=True)

def find_rooms_by_rows(name):
    room_matches = []
    for idx in range(len(data)):
        row = data.iloc[idx]
        if name in row['nameList']:
            room = row["roomNum"]
            room_matches.append(room)
        if len(room_matches) == 2:
            break
    return room_matches

for i in range (len(data)):
    name = data.loc[i,'name']
    room = find_rooms_by_rows(name)
    for j in range (len(room)):
        data.loc[i,'room'+str(j+1)] = room[j]

final_out_name = "dataSheet.xlsx"
final_df = data[["name", "sport", "room1", "room2", "food"]]
final_df.to_excel(final_out_name, index=False)
