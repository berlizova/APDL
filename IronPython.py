# The script generates an Excel file by extracting MODAL analysis results from the SOLVE.out file
# Author: Berlizova T. 09/11/2024
# analysis = ExtAPI.DataModel.Project.Model.Analyses[1] - the modal analysis number in the project tree
#
# The Excel file will be saved in the analysis folder at the specified path
#


import clr
import os
import csv
from System import Type
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
x_direction_data = [["MODE", "FREQUENCY", "RATIO X"]]
y_direction_data = [["RATIO Y"]]
z_direction_data = [["RATIO Z"]]
rot_x_direction_data = [["RATIO RotX"]]
rot_y_direction_data = [["RATIO RotY"]]
rot_z_direction_data = [["RATIO RotZ"]]
mass_summary_data = [["X-DIR", "RATIO% X", "Y-DIR", "RATIO% Y", "Z-DIR", "RATIO% Z"]]

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
            x_direction_data.append([columns[0], columns[1], columns[4]])
        elif in_y_direction:
            if columns[0] == 'MODE':
                continue
            y_direction_data.append([columns[4]])
        elif in_z_direction:
            if columns[0] == 'MODE':
                continue
            z_direction_data.append([columns[4]])
        elif in_rotx_direction:
            if columns[0] == 'MODE':
                continue
            rot_x_direction_data.append([columns[4]])
        elif in_roty_direction:
            if columns[0] == 'MODE':
                continue
            rot_y_direction_data.append([columns[4]])
        elif in_rotz_direction:
            if columns[0] == 'MODE':
                continue
            rot_z_direction_data.append([columns[4]])
        elif in_effective_mass_section and len(columns) >= 11:
            if columns[0] == 'MODE':
                continue
            mass_summary_data.append([columns[5], columns[6], columns[7], columns[8], columns[9], columns[10]])

# Initialize Excel
excel_app = Excel.ApplicationClass()
excel_app.Visible = False
workbook = excel_app.Workbooks.Add()
sheet = workbook.ActiveSheet

# Write the headers
headers = ["MODE", "FREQUENCY", "RATIO X", "RATIO Y", "RATIO Z", "RATIO RotX", "RATIO RotY", "RATIO RotZ",
           "X-DIR", "RATIO% X", "Y-DIR", "RATIO% Y", "Z-DIR", "RATIO% Z"]
for i, header in enumerate(headers, start=1):
    sheet.Cells(1, i).Value2 = header
    sheet.Cells(1, i).Font.Bold = True
    sheet.Cells(1, i).Interior.ColorIndex = 36

# Define green fill for conditional formatting
green_fill = Excel.XlRgbColor.rgbLightGreen

# Write data into Excel sheet
for row_idx, (x_row, y_row, z_row, rotx_row, roty_row, rotz_row, summary_row) in enumerate(zip(
        x_direction_data[1:], y_direction_data[1:], z_direction_data[1:],
        rot_x_direction_data[1:], rot_y_direction_data[1:], rot_z_direction_data[1:], mass_summary_data[1:]), start=2):

    highlight_mode_column = False  # Flag to check if we need to highlight the MODE column

    # Write MODE and FREQUENCY
    sheet.Cells(row_idx, 1).Value2 = x_row[0]  # MODE
    sheet.Cells(row_idx, 2).Value2 = x_row[1]  # FREQUENCY

    # Write ratios and apply formatting if needed
    for col, value in enumerate([x_row[2], y_row[0], z_row[0], rotx_row[0], roty_row[0], rotz_row[0]], start=3):
        cell = sheet.Cells(row_idx, col)
        cell.Value2 = float(value)
        if float(value) >= 1:
            cell.Interior.Color = green_fill
            highlight_mode_column = True

    # Write summary data
    for col, value in enumerate(summary_row, start=9):
        cell = sheet.Cells(row_idx, col)
        cell.Value2 = float(value)

    # Highlight MODE column if any cell was highlighted
    if highlight_mode_column:
        sheet.Cells(row_idx, 1).Interior.Color = green_fill

# Apply conditional formatting for maximum 2 values in each "RATIO%" column
for col in [10, 12, 14]:
    values = [sheet.Cells(row, col).Value2 for row in range(2, sheet.UsedRange.Rows.Count + 1)]
    top_two_values = sorted(values, reverse=True)[:2]
    for row in range(2, sheet.UsedRange.Rows.Count + 1):
        cell_value = sheet.Cells(row, col).Value2
        if cell_value in top_two_values:
            sheet.Cells(row, col).Interior.Color = green_fill

# Auto fit the columns
sheet.Columns.AutoFit()

# Save the workbook
workbook.SaveAs(output_excel_path)
workbook.Close(False)
excel_app.Quit()

# Release COM objects
Marshal.ReleaseComObject(sheet)
Marshal.ReleaseComObject(workbook)
Marshal.ReleaseComObject(excel_app)

print("Results extracted, saved, and formatted in", output_excel_path)
