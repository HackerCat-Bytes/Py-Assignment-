import pandas as pd
import numpy as np

# Read the data from first sheet into a Pandas DataFrame
frame1 = pd.read_excel("Python_Assignment.xlsx", sheet_name="(Input) User IDs", usecols="D:G", skiprows=10)
frame1.columns = ["S No", "Team Name", "total_statement", "total_reason"]
frame1.dropna(how="all", inplace=True)

# Read the data from second sheet into a Pandas DataFrame
frame2 = pd.read_excel("Python_Assignment.xlsx", sheet_name="(Input) Rigorbuilder RAW", usecols="C:G", skiprows=7)
frame2.dropna(how="all", inplace=True)
frame2.columns = ["S No", "Name", "User ID", "total_statements", "total_reasons"]

# Merge the two data frames into one
frame = pd.merge(left=frame1, right=frame2, on="S No")

# Convert columns to numeric type
frame["total_statement"] = pd.to_numeric(frame["total_statement"], errors="coerce")
frame["total_reason"] = pd.to_numeric(frame["total_reason"], errors="coerce")

# Calculate the average statements and reasons for each team
team_stats = frame.groupby(["Team Name"])[["total_statement", "total_reason"]].mean()
team_stats.columns = ["Average Statements", "Average Reasons"]
team_stats.sort_values(by=["Average Statements", "Average Reasons"], ascending=False, inplace=True)

# Calculate the individual ranks
ranks = frame[["Name", "User ID", "total_statements", "total_reasons"]].sort_values(by=["total_statements", "total_reasons"], ascending=False)
ranks.index.name = "Rank"

# Write the output to an Excel file
with pd.ExcelWriter("output.xlsx") as writer:
    team_stats.to_excel(writer, sheet_name="Sheet1")
    ranks.to_excel(writer, sheet_name="Sheet2")