import tkinter as tk
from tkinter import filedialog
from PIL import Image, ImageTk
import pyodbc
import pandas as pd
from datetime import datetime

def export_to_excel():
    database_path = r'./GEN.mdb'
    conn_str = f'DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={database_path};'
    conn = pyodbc.connect(conn_str)
    cursor = conn.cursor()
    table_name = 'FloatTable'

    start_date_str = start_date_var.get()
    end_date_str = end_date_var.get()

    start_date = datetime.strptime(start_date_str, "%Y-%m-%d %H:%M:%S")
    end_date = datetime.strptime(end_date_str, "%Y-%m-%d %H:%M:%S")

    query = f"SELECT * FROM {table_name} WHERE DateAndTime >= ? AND DateAndTime <= ?;"
    cursor.execute(query, start_date, end_date)
    rows = cursor.fetchall()
    conn.close()

    data_list = []
    for row in rows:
        data_list.append({
            'DateAndTime': row[0].strftime('%d-%m-%Y %H:%M:%S'),
            'Millitm': row[1],
            'TagIndex': row[2],
            'Val': row[3],
            'Status': row[4],
            'Marker': row[5]
        })

    df = pd.DataFrame(data_list)

    file_name = filedialog.asksaveasfilename(defaultextension=".xlsx", filetypes=[("Excel Files", "*.xlsx")])
    if file_name:
        df.to_excel(file_name, index=False)
        status_label.config(text=f"Data exported to {file_name} successfully.")

# Create the main application window
app = tk.Tk()
app.title("Data Export App")
app.geometry("600x400")  # Increased window size

# Create a frame to hold the logo and heading
header_frame = tk.Frame(app)
header_frame.pack(pady=10)

# Load and display the company logo (assuming it's in the same directory with the name "logo.png")
logo_image = Image.open("R.jpeg")
logo_image = logo_image.resize((50, 50), Image.BICUBIC)  # Increased logo size
logo_photo = ImageTk.PhotoImage(logo_image)
logo_label = tk.Label(app, image=logo_photo)
logo_label.pack(pady=10)  # Add padding to space it from the heading

# Create and display the heading
heading_label = tk.Label(app, text="Data Export App", font=("Helvetica", 20, "bold"))  # Increased font size
heading_label.pack(pady=10)  # Add padding to space it from the entry fields

# Create and display labels, entry fields, and buttons
start_label = tk.Label(app, text="Enter the start date and time (YYYY-MM-DD HH:MM:SS):")
start_label.pack()
start_date_var = tk.StringVar()
start_date_entry = tk.Entry(app, textvariable=start_date_var)
start_date_entry.pack()

end_label = tk.Label(app, text="Enter the end date and time (YYYY-MM-DD HH:MM:SS):")
end_label.pack()
end_date_var = tk.StringVar()
end_date_entry = tk.Entry(app, textvariable=end_date_var)
end_date_entry.pack()

export_button = tk.Button(app, text="Export Data", command=export_to_excel)
export_button.pack()

status_label = tk.Label(app, text="")
status_label.pack()

# Text in the bottom right corner
code_by_label = tk.Label(app, text="Code written by Radhika Bhoyar", font=("Helvetica", 10))
code_by_label.pack(side="bottom", anchor="se", padx=10, pady=10)

# Start the main event loop
app.mainloop()
