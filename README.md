# 📌 User Guide: Excel Conversion ➡️ KMZ

This script transforms an Excel file containing GPS coordinates into a KMZ file that can be used in Google Earth.

## 📂 Prerequisites

## 1️⃣ Install Python and dependencies

Make sure that Python is installed on your machine.

Install the necessary libraries with the following command:

`pip install pandas openpyxl`

## 2️⃣ Prepare the Excel file

The Excel file must contain a sheet with the following columns:

- CODE GDO DÉPART
- NOM DU POSTE (Name of the point displayed on the map)
- COMMUNE
- GPS X (Longitude in WGS84)
- GPS Y (Latitude in WGS84)
- DEPART

## 🛑 Attention: Coordinates must be in decimal degrees (e.g., 2.3522, 48.8566).

## 🚀 Using the script

## 1️⃣ Place your Excel file in the script folder

Edit the following lines in the script to specify the name of your Excel file and the sheet to use:

```python
excel_file = "file.xlsx"  # Name of your Excel file
sheet_name = "Feuil1"  # Excel sheet name
```

## 2️⃣ Run the script

In a terminal or command prompt, type:

`python script.py`

The script will:

- Load data from the Excel file
- Check that all necessary columns are present
- Generate a `points.kmz` file containing the geolocated points with the associated information

## 3️⃣ Open KMZ file in Google Earth

Once `points.kmz` is generated, open Google Earth Pro and then:

- Go to **File > Open**
- Select `points.kmz`

All points will be displayed on the map with their descriptions 🎉

## 🔍 Troubleshooting

❌ **Error: Missing columns**

If the script displays an error indicating that some columns are missing, check that your Excel file contains all the required columns.

❌ **KMZ file does not display correctly**

- Check that the GPS coordinates are in decimal format (e.g., `2.35` and not `2°35'`).
- Make sure that your Excel cells do not contain any spaces or special characters.

## 📌 Conclusion

This script allows you to easily convert your Excel files to KMZ, making it easy to view them in Google Earth. Feel free to adapt the script if necessary! 🚀
