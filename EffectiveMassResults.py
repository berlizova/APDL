# The script generates an Excel file by extracting MODAL analysis results from the SOLVE.out file
# Author: t.berlizova.ext 09/11/2024
# analysis = ExtAPI.DataModel.Project.Model.Analyses[1] - the modal analysis number in the project tree
#
# The Excel file will be saved in the analysis folder at the specified path
#
# -*- coding: utf-8 -*-
import clr
import os
from System.Runtime.InteropServices import Marshal

# Import Excel Interop
clr.AddReference("Microsoft.Office.Interop.Excel")
from Microsoft.Office.Interop import Excel

# Get the path to solve.out using ANSYS working directory
analysis = ExtAPI.DataModel.Project.Model.Analyses[1]  # Adjust the analysis index
solve_out_path = analysis.WorkingDir + "\\solve.out"

# Set the path to the output Excel file
output_excel_path = os.path.join(os.path.dirname(solve_out_path), "EffectiveMassResults.xlsx")

# Initialize data structures
x_direction_data = []
y_direction_data = []
z_direction_data = []
rot_x_direction_data = []
rot_y_direction_data = []
rot_z_direction_data = []
mass_summary_data = []

# Read the solve.out file
with open(solve_out_path, 'r') as file:
    lines = file.readlines()

# Flags for section identification
in_x_direction = False
in_y_direction = False
in_z_direction = False
in_rotx_direction = False
in_roty_direction = False
in_rotz_direction = False
in_effective_mass_section = False

# Process each line in the file
for line in lines:
    if "PARTICIPATION FACTOR CALCULATION *****  X  DIRECTION" in line:
        in_x_direction = True
        in_y_direction = in_z_direction = in_rotx_direction = in_roty_direction = in_rotz_direction = in_effective_mass_section = False
        continue
    elif "PARTICIPATION FACTOR CALCULATION *****  Y  DIRECTION" in line:
        in_y_direction = True
        in_x_direction = in_z_direction = in_rotx_direction = in_roty_direction = in_rotz_direction = in_effective_mass_section = False
        continue
    elif "PARTICIPATION FACTOR CALCULATION *****  Z  DIRECTION" in line:
        in_z_direction = True
        in_x_direction = in_y_direction = in_rotx_direction = in_roty_direction = in_effective_mass_section = False
        continue
    elif "PARTICIPATION FACTOR CALCULATION *****ROTX DIRECTION" in line:
        in_rotx_direction = True
        in_x_direction = in_y_direction = in_z_direction = in_roty_direction = in_effective_mass_section = False
        continue
    elif "PARTICIPATION FACTOR CALCULATION *****ROTY DIRECTION" in line:
        in_roty_direction = True
        in_x_direction = in_y_direction = in_z_direction = in_rotx_direction = in_effective_mass_section = False
        continue
    elif "PARTICIPATION FACTOR CALCULATION *****ROTZ DIRECTION" in line:
        in_rotz_direction = True
        in_x_direction = in_y_direction = in_z_direction = in_rotx_direction = in_roty_direction = in_effective_mass_section = False
        continue
    elif "***** MODAL MASSES, KINETIC ENERGIES, AND TRANSLATIONAL EFFECTIVE MASSES SUMMARY *****" in line:
        in_effective_mass_section = True
        in_x_direction = in_y_direction = in_z_direction = in_rotx_direction = in_roty_direction = False
        continue

    if "sum" in line:
        in_x_direction = in_y_direction = in_z_direction = in_rotx_direction = in_roty_direction = in_rotz_direction = in_effective_mass_section = False

    columns = line.split()
    if len(columns) >= 8:
        if in_x_direction:
            if columns[0] == 'MODE':
                continue
            x_direction_data.append([columns[0], columns[1], columns[4], columns[5]])
        elif in_y_direction:
            if columns[0] == 'MODE':
                continue
            y_direction_data.append([columns[0], columns[1], columns[4], columns[5]])
        elif in_z_direction:
            if columns[0] == 'MODE':
                continue
            z_direction_data.append([columns[0], columns[1], columns[4], columns[5]])
        elif in_rotx_direction:
            if columns[0] == 'MODE':
                continue
            rot_x_direction_data.append([columns[0], columns[1], columns[4], columns[5]])
        elif in_roty_direction:
            if columns[0] == 'MODE':
                continue
            rot_y_direction_data.append([columns[0], columns[1], columns[4], columns[5]])
        elif in_rotz_direction:
            if columns[0] == 'MODE':
                continue
            rot_z_direction_data.append([columns[0], columns[1], columns[4], columns[5]])
        elif in_effective_mass_section and len(columns) >= 11:
            if columns[0] == 'MODE':
                continue
            mass_summary_data.append(
                [columns[0], columns[1], columns[2], columns[5], columns[6], columns[7], columns[8],
                 columns[9], columns[10]])

# Initialize Excel
excel_app = Excel.ApplicationClass()
excel_app.Visible = False
workbook = excel_app.Workbooks.Add()

# Define the green fill color for highlighting
green_fill = Excel.XlRgbColor.rgbLightGreen

# Define yellow fill color for the headers
yellow_fill = Excel.XlRgbColor.rgbYellow


# Function to write and format direction sheets
def write_and_format_direction_sheet(sheet, data, headers=None):
    if headers:
        for i, header in enumerate(headers, start=1):
            sheet.Cells(1, i).Value2 = header
            sheet.Cells(1, i).Font.Bold = True
            sheet.Cells(1, i).Interior.Color = yellow_fill  # Apply yellow fill to the headers
        start_row = 2
    else:
        start_row = 1

    # Write data
    for row_idx, row in enumerate(data, start=start_row):
        highlight_mode_column = False  # Flag to check if we need to highlight the MODE column
        for col_idx, cell_value in enumerate(row, start=1):
            cell = sheet.Cells(row_idx, col_idx)
            cell.Value2 = cell_value

            # Formatting only to the 3rd column (RATIO)
            if col_idx == 3:
                try:
                    value = float(cell_value)
                    if value >= 0.5:
                        cell.Interior.Color = green_fill
                        highlight_mode_column = True
                except ValueError:
                    pass

        # Highlight MODE column if any cell in the row was highlighted
        if highlight_mode_column:
            sheet.Cells(row_idx, 1).Interior.Color = green_fill

    sheet.Columns.AutoFit()


# Write direction sheets
def write_sheet_by_name(sheet_name, data):
    sheet = workbook.Sheets.Add()
    sheet.Name = sheet_name
    write_and_format_direction_sheet(sheet, data, headers=["MODE No", "FREQUENCY, Hz", "RATIO", "Effective Mass, ton"])


# Write direction sheets
write_sheet_by_name("X Direction", x_direction_data)
write_sheet_by_name("Y Direction", y_direction_data)
write_sheet_by_name("Z Direction", z_direction_data)
write_sheet_by_name("RotX Direction", rot_x_direction_data)
write_sheet_by_name("RotY Direction", rot_y_direction_data)
write_sheet_by_name("RotZ Direction", rot_z_direction_data)

# Create and write to the Effective Mass sheet
mass_summary_sheet = workbook.Sheets.Add()
mass_summary_sheet.Name = "Effective Mass"
write_and_format_direction_sheet(mass_summary_sheet, mass_summary_data,
                                 headers=["MODE No", "FREQUENCY, Hz", "MODAL MASS", "X-DIR", "RATIO%", "Y-DIR",
                                          "RATIO%", "Z-DIR", "RATIO%"])


# Highlight top 6 values in columns 5, 7, and 9 for Effective Mass
def highlight_top_values(sheet, columns, top_n=6):
    for col in columns:
        values = [sheet.Cells(row, col).Value2 for row in range(2, sheet.UsedRange.Rows.Count + 1)]
        top_values = sorted(values, reverse=True)[:top_n]
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            cell_value = sheet.Cells(row, col).Value2
            if cell_value in top_values:
                sheet.Cells(row, col).Interior.Color = green_fill
                sheet.Cells(row, 1).Interior.Color = green_fill


highlight_top_values(mass_summary_sheet, [5, 7, 9])

# Create and write to the summary sheet
summary_sheet = workbook.Sheets.Add(Before=workbook.Sheets(1))
summary_sheet.Name = "Summary"
summary_data = [["MODE No", "FREQUENCY", "X Direction Ratio", "Y Direction Ratio", "Z Direction Ratio", "Rot X Ratio",
                 "Rot Y Ratio", "Rot Z Ratio", "X-DIR", "RATIO% X", "Y-DIR", "RATIO% Y", "Z-DIR", "RATIO% Z"]]

for i in range(0, len(x_direction_data)):
    row = [
        x_direction_data[i][0],
        x_direction_data[i][1],
        x_direction_data[i][2],
        y_direction_data[i][2],
        z_direction_data[i][2],
        rot_x_direction_data[i][2],
        rot_y_direction_data[i][2],
        rot_z_direction_data[i][2],
        mass_summary_data[i][3],
        mass_summary_data[i][4],
        mass_summary_data[i][5],
        mass_summary_data[i][6],
        mass_summary_data[i][7],
        mass_summary_data[i][8],
    ]
    summary_data.append(row)

# Write headers for summary sheet
for i, header in enumerate(summary_data[0], start=1):
    summary_sheet.Cells(1, i).Value2 = header
    summary_sheet.Cells(1, i).Font.Bold = True
    summary_sheet.Cells(1, i).Interior.Color = yellow_fill  # Apply yellow fill to the headers

# Write data to summary sheet starting from row 2
for i, row in enumerate(summary_data[1:], start=2):
    for j, cell_value in enumerate(row, start=1):
        summary_sheet.Cells(i, j).Value2 = cell_value


# Formatting for Summary sheet
def format_summary_sheet(sheet):
    # Highlight cells where columns from 3 to 8 have values >= 1
    for row in range(2, sheet.UsedRange.Rows.Count + 1):
        highlight_mode_column = False
        for col in range(3, 9):
            try:
                if float(sheet.Cells(row, col).Value2) >= 1:
                    sheet.Cells(row, col).Interior.Color = green_fill
                    highlight_mode_column = True
            except ValueError:
                pass

        # Highlight the MODE cell in column 1 if any cell in columns 3-8 was highlighted
        if highlight_mode_column:
            sheet.Cells(row, 1).Interior.Color = green_fill

    # Highlight top 1 value in columns 10, 12, and 14 (RATIO% X, RATIO% Y, RATIO% Z)
    for col in [10, 12, 14]:
        values = [sheet.Cells(row, col).Value2 for row in range(2, sheet.UsedRange.Rows.Count + 1)]
        top_value = max(values)
        for row in range(2, sheet.UsedRange.Rows.Count + 1):
            if sheet.Cells(row, col).Value2 == top_value:
                sheet.Cells(row, col).Interior.Color = green_fill
                sheet.Cells(row, 1).Interior.Color = green_fill


# Apply formatting to the summary sheet
format_summary_sheet(summary_sheet)

summary_sheet.Columns.AutoFit()

# Save the workbook
workbook.SaveAs(output_excel_path)
workbook.Close(False)
excel_app.Quit()

# Release COM objects
Marshal.ReleaseComObject(workbook)
Marshal.ReleaseComObject(excel_app)

print("Results extracted and saved into separate sheets in", output_excel_path)
