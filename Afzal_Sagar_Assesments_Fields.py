import arcpy
import os

gdb_path = r"I:\PMU\NEOGEO COLONIES\AFZAL SAGAR\AFZAL SAGAR\AFZAL SAGAR.gdb"

target_name = "Assesments"

found_fc = None

print("=" * 100)
print("Searching for:", target_name)
print("=" * 100)

# Search entire GDB
for root, datasets, feature_classes in arcpy.da.Walk(
    gdb_path,
    datatype="FeatureClass"
):
    for fc in feature_classes:

        if fc.lower() == target_name.lower():
            found_fc = os.path.join(root, fc)
            break

    if found_fc:
        break

if not found_fc:

    print("\nERROR: Assesments was not found in the GDB.")

else:

    print("\nFOUND:")
    print(found_fc)

    print("\n" + "=" * 120)
    print("FIELDS IN ASSESMENTS")
    print("=" * 120)

    fields = arcpy.ListFields(found_fc)

    print(
        f"{'No.':<5}"
        f"{'Field Name':<40}"
        f"{'Alias':<40}"
        f"{'Type':<20}"
        f"{'Length':<10}"
    )

    print("-" * 115)

    for i, field in enumerate(fields, 1):

        print(
            f"{i:<5}"
            f"{field.name:<40}"
            f"{field.aliasName:<40}"
            f"{field.type:<20}"
            f"{field.length:<10}"
        )

    print("-" * 115)
    print(f"Total Fields: {len(fields)}")