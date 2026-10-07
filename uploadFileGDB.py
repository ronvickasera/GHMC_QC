import arcpy
import os

# File Geodatabase path
#gdb_path = r"I:\PMU\NEOGEO COLONIES\ADARSH NAGAR\ADARSH NAGAR\ADARSH NAGAR.gdb"
#gdb_path = r"I:\PMU\NEOGEO COLONIES\AFZAL SAGAR\AFZAL SAGAR\AFZAL SAGAR.gdb"
gdb_path =  r"I:\PMU\NEOGEO COLONIES\AHMED NAGAR\AHMED NAGAR\AHMED NAGAR.gdb"

# Get current ArcGIS Pro project and active map
aprx = arcpy.mp.ArcGISProject("CURRENT")
m = aprx.activeMap

if m is None:
    raise RuntimeError("Please open a map in ArcGIS Pro.")

# Check if GDB exists
if not arcpy.Exists(gdb_path):
    raise FileNotFoundError(f"GDB not found: {gdb_path}")

# Add GDB connection to ArcGIS Pro project
aprx.updateDatabases(
    [{
        "databasePath": gdb_path,
        "isDefaultDatabase": False
    }],
    validate=False
) if False else None

# Set workspace
arcpy.env.workspace = gdb_path
          
# Add all feature classes, including those in feature datasets
datasets = [""] + (arcpy.ListDatasets("", "Feature") or [])

for dataset in datasets:
    workspace = os.path.join(gdb_path, dataset)

    arcpy.env.workspace = workspace

    for fc in arcpy.ListFeatureClasses() or []:
        fc_path = os.path.join(workspace, fc)

        try:
            m.addDataFromPath(fc_path)
            print(f"Added: {fc}")
        except Exception as e:
            print(f"Failed: {fc} - {e}")

# Add standalone tables
arcpy.env.workspace = gdb_path

for table in arcpy.ListTables() or []:
    table_path = os.path.join(gdb_path, table)

    try:
        m.addDataFromPath(table_path)
        print(f"Added table: {table}")
    except Exception as e:
        print(f"Failed table: {table} - {e}")

# Save project
aprx.save()

print("All feature classes and tables added successfully.")