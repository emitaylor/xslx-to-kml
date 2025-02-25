# 📌 User Guide: Excel Conversion ➡️ KML

This script transforms an Excel file containing GPS coordinates into a KML file that can be used in Google Earth.

## 📂 Prerequisites

## 1️⃣ Install Python and dependencies

Make sure that Python is installed on your machine.

Install the necessary libraries with the following command:

`pip install pandas openpyxl`

## 2️⃣ Prepare the Excel file

The Excel file must contain a sheet with the following columns:

- Sector

- GDO

- Name (Name of the point displayed on the map)

- Commune

- GPS X Start (Longitude in WGS84)

- GPS Y Start (Latitude in WGS84)

- Departure

## 🛑 Attention: Coordinates must be in decimal degrees (ex: 2.3522, 48.8566).

## 🚀 Using the script

## 1️⃣ Place your Excel file in the script folder

Edit the following line in the script to specify the name of your Excel file and the sheet to use:

excel_file = "PDT.xlsx"  # Name of your Excel file
sheet_name = "Vague Test 1"  # Excel tab name

## 2️⃣ Run the script

In a terminal or command prompt, type:

`python script.py`

The script will:

Load data from Excel file

Check that all necessary columns are present

Generate a points.kml file containing the geolocated points with the associated information

## 3️⃣ Open KML file in Google Earth

Once points.kml is generated, open Google Earth Pro and then:

Go to File Open

Select points.kml

All points will be displayed on the map with their descriptions 🎉

## 🔍 Troubleshooting

❌ Error: Missing columns

If the script displays an error that some columns are missing, check that your Excel file contains all the required columns.

❌ KML file does not display correctly

Check that the GPS coordinates are in decimal format (ex: 2.35 and not 2°35').

Make sure that your Excel cells do not contain any spaces or special characters.

## 📌 Conclusion

This script allows you to easily convert your Excel files to KML, making it easy to view them in Google Earth. Do not hesitate to adapt the script if necessary! 🚀