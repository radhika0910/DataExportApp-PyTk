# Data Export App with Python and Tkinter

This Python application allows users to export data from a Microsoft Access database into an Excel file. The user can specify a start and end date range to select the data they want to export. The data is then retrieved from the Access database using the Pyodbc library and converted into a pandas DataFrame. Finally, the selected data is exported to an Excel file using the Pandas library.

## Features:
- User-friendly GUI built with Tkinter for easy data selection.
- Ability to specify a date and time range to filter data from the database.
- Seamless integration with the Microsoft Access database using Pyodbc.
- Data manipulation and export to Excel using the powerful Pandas library.

## Customizations:
- Increased window size and added a company logo for a visually appealing interface.
- Displayed the code creator's name (Radhika Bhoyar) in the bottom right corner.
- Improved font size and style for better readability.

## Dependencies:
- Python 3.x
- Tkinter (included in the standard library)
- Pandas (install using 'pip install pandas')
- Pyodbc (install using 'pip install pyodbc')
- Pillow (install using 'pip install pillow') for image manipulation.

## Usage:
1. Ensure you have the necessary dependencies installed.
2. Clone this repository to your local machine.
3. Replace the 'R.png' file with your own company logo in the same directory.
4. Run the 'data_export_app.py' script.
5. Enter the start and end date range to select the data.
6. Click the 'Export Data' button to save the data as an Excel file.

## Contributions:
Contributions to enhance this Data Export App are welcome! If you encounter any issues, feel free to open an issue in this repository.

## Author:
This Data Export App is written by Radhika Bhoyar.
