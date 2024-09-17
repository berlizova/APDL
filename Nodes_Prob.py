# This script generates probes for specific nodes.
# It can create probes for different nodes and results.
# To run the script, an Excel table with the node numbers of interest is required.
# Nodes for different results must be placed on separate sheets in the Excel file.

import clr

clr.AddReference('Microsoft.Office.Interop.Excel')
from Microsoft.Office.Interop import Excel


# Function to read nodeIDs from a specific Excel sheet using COM
def read_node_ids_from_excel(file_path, sheet_number):
    excel_app = Excel.ApplicationClass()
    workbook = excel_app.Workbooks.Open(file_path)

    try:
        # Access the sheet by its number (1-based index)
        sheet = workbook.Sheets(sheet_number)
    except Exception as e:
        print("Error accessing sheet number " + str(sheet_number) + ": " + str(e))
        workbook.Close(False)
        excel_app.Quit()
        return []

    node_ids = []
    row = 2  # Starting from row 2, assuming row 1 contains headers
    while True:
        node_id = sheet.Cells(row, 1).Value2  # Assuming nodeID is in the first column
        if node_id is None:
            break
        node_ids.append(int(node_id))
        row += 1

    workbook.Close(False)
    excel_app.Quit()
    return node_ids


##################################################################################################################
# Section where changes need to be made
###################################################################################################################
# Start
# Specify the path to the Excel file
excel_file_path = r'K:\Groups\ONGDSGE3\BE\RESS\000_INITIAL_STUDIES\04_QAD_3\04_ANALYSIS\t.berlizova\002_script\001_Nod_ID_Script\Nodes.xlsx'

# Specify the analysis (change the index as needed)
analysis = ExtAPI.DataModel.Project.Model.Analyses[2]  # [0] - Analysis number

# Define result objects and corresponding sheet numbers in the Excel file
result_objects_sheets = {

    'Main_beam': 1,  # 'Name Results': Sheet number 1
    'Badframe': 2,  # 'Name Results': Sheet number 2
    'Gen_Rear_beam': 4,  # 'Name Results': Sheet number 4
    'Gen_Front_beam': 3,  # 'Name Results': Sheet number 3
    'Front_beam': 5,  # 'Name Results': Sheet number 1
    'Rear_beam': 6,  # 'Name Results': Sheet number 1
    'CT_inetface': 7,  # 'Name Results': Sheet number 1
    'TU_interface': 8,  # 'Name Results': Sheet number 1
    'Total Deformations (RESS+NSS only)': 9,

    # Add more results and corresponding sheet numbers as needed
}

# Finish
###################################################################################################################

###################################################################################################################
# Loop through the defined result objects and corresponding sheet numbers
###################################################################################################################
for result_name, sheet_number in result_objects_sheets.items():
    # Get all result objects from the selected analysis
    results_in_analysis = analysis.Solution.Children  # Retrieve all results for the selected analysis

    # Check if there is a result with the specified name in the analysis
    resultObjects = [res for res in results_in_analysis if res.Name == result_name]

    # Check if the object with the given name exists
    if resultObjects is None or len(resultObjects) == 0:
        print("No object found with name: " + result_name + " in the specified analysis.")
        continue  # Skip to the next object if none found

    resultObject = resultObjects[0]  # Take the first matching object

    # Read nodeIDs from the corresponding sheet number
    node_ids = read_node_ids_from_excel(excel_file_path, sheet_number)

    # Create probes for all nodeIDs from the sheet
    for nodeID in node_ids:
        probeLabel = Graphics.LabelManager.CreateProbeLabel(resultObject)
        probeLabel.Scoping.Node = nodeID
        # Print the message including the result name, node ID, and analysis name
        print("Probe created for Node " + str(nodeID) + " in " + result_name + " (Analysis: " + analysis.Name + ")")
