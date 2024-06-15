import tkinter as tk
from tkinter import ttk

# Create the main window
root = tk.Tk()
root.title("Financial Analysis Tool")
root.geometry("800x600")  # Set the size of the window

# Create the tab control widget
tab_control = ttk.Notebook(root)

# Create the tabs
tab1 = ttk.Frame(tab_control)
tab2 = ttk.Frame(tab_control)

# Add tabs to the notebook
tab_control.add(tab1, text='Data Analysis')
tab_control.add(tab2, text='Results')

tab_control.pack(expand=1, fill="both")

# Run the main loop
root.mainloop()
