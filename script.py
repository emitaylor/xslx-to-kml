import pandas as pd
from xml.etree.ElementTree import Element, SubElement, ElementTree

excel_file = "file.xlsx" # excel file name
sheet_name = "Feuil1"    # sheet name
kml_file = "points.kml"  # futur kml name

# load excel file and sheet
df = pd.read_excel(excel_file, sheet_name=sheet_name)

# find if columns exists ? 
required_columns = {"CODE GDO DÉPART", "NOM DU POSTE", "COMMUNE", "GPS X", "GPS Y", "DEPART"}
if not required_columns.issubset(df.columns):
    raise ValueError(f"Required columns are missing. Add: {required_columns - set(df.columns)}")

# create KML
kml = Element("kml", xmlns="http://www.opengis.net/kml/2.2")
document = SubElement(kml, "Document")

# add KML points
for _, row in df.iterrows():
    placemark = SubElement(document, "Placemark")
    
    # point name
    SubElement(placemark, "name").text = str(row["NOM DU POSTE"])

    # Description generated from columns
    description = f"""
    GDO : {row["CODE GDO DÉPART"]}
    Commune : {row["COMMUNE"]}
    Départ : {row["DEPART"]}
    """
    SubElement(placemark, "description").text = description.strip()

    # GPS coordinates (Longitude, Latitude)
    point = SubElement(placemark, "Point")
    SubElement(point, "coordinates").text = f'{row["GPS X"]},{row["GPS Y"]},0'

# Save the KML
tree = ElementTree(kml)
tree.write(kml_file, encoding="utf-8", xml_declaration=True)

print(f"✅ Generated KML file: {kml_file}")
