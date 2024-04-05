
import tkinter as tk
import convert_xls2xlsx
import create_data_summary_MFP126_MFP166_MFP270

def button_click():
    user_input = entry.get()
    file_path = "C:/Users/Noly.Espardinez/Desktop/MFP00166/python/location.txt"  # Specify your desired file path here
    with open(file_path, "w") as file:
        file.write(user_input)
    print("Input text saved to:", file_path)
    root.destroy()

# Create the main window
root = tk.Tk()
root.title("MFP00126 & MFP00166")

bold_font = ("Helvetica", 12, "bold")

# Create a label with bold text
label = tk.Label(root, text="Result Logs Location:", font=bold_font)
label.pack()

# Create an entry box for input
entry = tk.Entry(root, width=50)  # Set width to 50 characters
entry.pack()

button = tk.Button(root, text="Submit", command=button_click, width=15)  # Set width to 20 characters
button.pack(pady=(3, 3))  # 2mm margin on top and bottom

# Run the Tkinter event loop
root.mainloop()

if __name__ == "__main__":
    # Assuming the text file contains the path
    with open('location.txt', 'r') as file:
        path = file.readline().strip()
    
    convert_xls2xlsx.xls_2_xlsx(path)
    create_data_summary_MFP126_MFP166_MFP270.processResults(path)