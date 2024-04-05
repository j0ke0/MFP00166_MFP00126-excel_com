import tkinter as tk
from PIL import Image, ImageTk
import convert_xls2xlsx
import create_data_summary_MFP126_MFP166_MFP270

def button_click():
    user_input = entry.get()
    file_path = "C:/Users/Noly.Espardinez/Desktop/MFP00166/python/Support/location.txt"  # Specify your desired file path here
    with open(file_path, "w") as file:
        file.write(user_input)
    print("Input text saved to:", file_path)

def remove_placeholder(event):
    if entry.get() == placeholder_text:
        entry.delete(0, tk.END)
        entry.config(fg='black')  # Change text color to black when placeholder is removed

def add_placeholder(event):
    if not entry.get():
        entry.insert(0, placeholder_text)
        entry.config(fg='grey')  # Change text color to grey for placeholder

# Create the main window
root = tk.Tk()
root.title("MFP00126 & MFP00166")

bold_font = ("Arial", 10,)

# Set the background color to grey
background_color = "#d9d9d9"
root.configure(bg=background_color)

# Load and resize the logo image
logo_image = Image.open("C:/Users/Noly.Espardinez/Desktop/MFP00166/python/Support/my_logo.png")  # Specify the path to your logo image
logo_image = logo_image.resize((300, 55), Image.LANCZOS)  # Use Image.LANCZOS for resizing
logo_photo = ImageTk.PhotoImage(logo_image)

# Create a label to display the logo
logo_label = tk.Label(root, image=logo_photo, bg=background_color)
logo_label.pack(anchor='n', padx=10, pady=10)  # Place the logo on the top-right corner with padding

# Create a label with bold text
label = tk.Label(root, text="This application will combine all \n excel results files into a single excel file:", font=bold_font, bg=background_color)
label.pack()

# Create an entry box for input
placeholder_text = "Enter folder location..."
entry = tk.Entry(root, width=50, borderwidth=2, relief="flat", fg='grey')  # Set width to 50 characters
entry.insert(0, placeholder_text)
entry.bind("<FocusIn>", remove_placeholder)
entry.bind("<FocusOut>", add_placeholder)
entry.pack(padx=(5, 5), pady=(5, 5))

def process_data():
    # Assuming the text file contains the path
    with open('C:/Users/Noly.Espardinez/Desktop/MFP00166/python/Support/location.txt', 'r') as file:
        path = file.readline().strip()

    convert_xls2xlsx.xls_2_xlsx(path)
    create_data_summary_MFP126_MFP166_MFP270.processResults(path)
    root.destroy()

# Button to save location
button = tk.Button(root, text="Save Location", command=button_click, width=13)  # Set width to 20 characters
button.pack(side=tk.LEFT, padx=(5, 5), pady=(5, 5))  # Position on the left with a 3mm margin on left and right, and a 3mm margin on top and bottom

# Button to process data
process_button = tk.Button(root, text="Process Results", command=process_data, width=13)
process_button.pack(side=tk.RIGHT, padx=(5, 5), pady=(5, 5))  # Position on the right with a 3mm margin on left and right, and a 3mm margin on top and bottom

# Run the Tkinter event loop
root.mainloop()
