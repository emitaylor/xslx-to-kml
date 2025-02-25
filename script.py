import pandas as pd
from xml.etree.ElementTree import Element, SubElement, ElementTree
import zipfile
import os

excel_file = "file.xlsx"  # Excel file name
sheet_name = "Feuil1"     # Sheet name
kml_file = "points.kml"   # Temporary KML file name
kmz_file = "points.kmz"   # Final KMZ file name

# Load the Excel file
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# Check if required columns exist
required_columns = {"CODE GDO DÉPART", "NOM DU POSTE", "COMMUNE", "GPS X", "GPS Y", "DEPART"}
if not required_columns.issubset(df.columns):
    raise ValueError(f"Required columns are missing. Add: {required_columns - set(df.columns)}")

# Create the KML file
kml = Element("kml", xmlns="http://www.opengis.net/kml/2.2")
document = SubElement(kml, "Document")

for _, row in df.iterrows():
    placemark = SubElement(document, "Placemark")
    
    # Point name
    SubElement(placemark, "name").text = str(row["NOM DU POSTE"])
    
    # Description with additional information
    description = f"""
    GDO : {row["CODE GDO DÉPART"]}
    Commune : {row["COMMUNE"]}
    Départ : {row["DEPART"]}
    """
    SubElement(placemark, "description").text = description.strip()
    
    # GPS coordinates (Longitude, Latitude)
    point = SubElement(placemark, "Point")
    SubElement(point, "coordinates").text = f'{row["GPS X"]},{row["GPS Y"]},0'

# Save the temporary KML file
tree = ElementTree(kml)
tree.write(kml_file, encoding="utf-8", xml_declaration=True)

# Create the KMZ file (ZIP compression of the KML file)
with zipfile.ZipFile(kmz_file, "w", zipfile.ZIP_DEFLATED) as kmz:
    kmz.write(kml_file, os.path.basename(kml_file))

# Remove the temporary KML file
os.remove(kml_file)

print(f"✅ KMZ file generated: {kmz_file}")