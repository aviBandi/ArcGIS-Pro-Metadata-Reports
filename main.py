"""
Author: Avinash Bandi
Title: ReviewDB_MetadataReport
Created: 7/23/2024

Purpose: To create an metadata report card once a year to be assigned as a helpdesk ticket to GIS Staff

Details:
    1. Create multi sheet xlsx spreadsheet
    2. Loop through each database. Get all tables, feature classes and raster
    3. Loop through each dataset and create dictionary list with all metadata details
    4. Goes through each value and sees if any are None, if so, it will be marked incomplete and break out of loop
    5. Create pandas dataframe using dictionary list of metadata


Sources:
    1. https://www.geeksforgeeks.org/how-to-write-pandas-dataframes-to-multiple-excel-sheets/
"""

# YOU NEED TO IMPORT YOUR OWN EMAIL SENDER MODULE
from Modules import Email

try:
    
    import arcpy
    from arcpy import metadata as md
    import pandas as pd
    from datetime import datetime
    
except Exception as e:
    importError_solution = 'As a solution please remote into APP-GISSCPT-P01 as SVC-GIS-SCRIPTS and log into ArcGIS' \
                           ' Pro as SVC-GIS-SCRIPTS. Another issue may be that pro was left open on this server.'
                           
    txt = 'SA - ReviewDB_MetadataReport.py Failed to import. ' + importError_solution
    subject = 'SA - ReviewDB_MetadataReport.py Failed to Import'
    Email.helpdesk(txt, subject, e)
    exit()

try:
    # You need to add your own arcy workspace
    arcpy.env.workspace = r""
    databases = arcpy.ListWorkspaces(workspace_type="SDE")
    date = datetime.today().strftime('%Y%m%d')
    # You need to add your own path
    xl_path = r""

    # Create multi sheet xlsx spreadsheet Source #1
    with pd.ExcelWriter(xl_path) as writer:
        # Loop through each database. Get all tables, feature classes and raster
        for database in databases:
            listOfMetaData = []
            arcpy.env.workspace = database
            dataList = arcpy.ListTables() + arcpy.ListFeatureClasses() + arcpy.ListRasters()
            databaseName = arcpy.env.workspace.split("\\")[-1].strip(".sde")

            # Loop through each dataset and create dictionary list with all metadata details
            for eachDataInList in dataList:
                dictMetaData = {"Database": None, "Dataset": None, "Tags": None, "Description": None, "Summary": None, "Credits": None, "Use Limitations": None, "Status": ""}
                dataListPath = arcpy.env.workspace + "\\" + eachDataInList
                dataItemMetaData = md.Metadata(dataListPath)

                dictMetaData["Database"] = databaseName
                dictMetaData["Dataset"] = eachDataInList
                dictMetaData["Summary"] = dataItemMetaData.summary
                dictMetaData["Credits"] = dataItemMetaData.credits
                dictMetaData["Tags"] = dataItemMetaData.tags
                dictMetaData["Use Limitations"] = dataItemMetaData.accessConstraints
                dictMetaData["Description"] = dataItemMetaData.description

                if dictMetaData["Use Limitations"] != None:
                    dictMetaData["Use Limitations"] = len(dataItemMetaData.accessConstraints)

                if dictMetaData["Description"] != None:
                    dictMetaData["Description"] = len(dataItemMetaData.description)

                # Goes through each value and sees if any are None, if so, it will be marked incomplete and break out of loop
                for eachValue in dictMetaData.values():
                    if eachValue == None:
                        dictMetaData["Status"] = "Incomplete"
                        break
                    dictMetaData["Status"] = "Complete"
                print("--------------------------------")
                print(f"{eachDataInList} complete in {database}")
                listOfMetaData.append(dictMetaData)

            # Create pandas dataframe using dictionary list of metadata
            df = pd.DataFrame(listOfMetaData)
            df.to_excel(writer, sheet_name=databaseName)

    # Add your own descriptions
    txt = ""
    subject = "Yearly Metadata Report"
    recipient = ""

    # Use your own custom email sender module
    Email.customGIS_attach(custom_txt=txt, subject=subject, recipient=recipient, attachment_path=xl_path,
                           content_type='plain')

except Exception as e:
    txt = 'SA - ReviewDB_MetadataReport.py Failed to Create xlsx. '
    subject = 'SA - ReviewDB_MetadataReport.py Failed to Create xlsx.'
    # Import and use your own custom email sender module.
    Email.helpdesk(txt, subject, e)