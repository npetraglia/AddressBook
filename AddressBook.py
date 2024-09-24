"""
Project Title: Address Book
Author: <Nick Petraglia>
Version: 1.1.0

Version Info: This updated version of the program allows for search and filter capability.

Description: The purpose of this program is to create an address book for important
Fidelity BSG contacts.
"""
from office365.sharepoint.client_context import ClientContext
from office365.runtime.auth.user_credential import UserCredential
import json
import os
import csv
import tkinter as tk
from tkinter import messagebox, ttk

# File where contacts will be saved
ADDRESS_BOOK_FILE = 'address_book.json'

# Load contacts from the file
def load_contacts():
    if os.path.exists(ADDRESS_BOOK_FILE):
        with open(ADDRESS_BOOK_FILE, 'r') as file:
            return json.load(file)
    return {}

# Save contacts to the file
def save_contacts():
    with open(ADDRESS_BOOK_FILE, 'w') as file:
        json.dump(address_book, file, indent=4)

# Add or update a contact
def add_or_update_contact():
    name = entry_name.get()
    phone = entry_phone.get()
    email = entry_email.get()
    location = entry_location.get()

    if name and phone and email and location:
        address_book[name] = {
            'Phone': phone,
            'Email': email,
            'Location': location
        }
        save_contacts()
        messagebox.showinfo("Success", f"Contact {name} added/updated successfully!")
        clear_entries()
        view_contacts()
    else:
        messagebox.showwarning("Input Error", "All fields are required.")

# Delete selected contact
def delete_contact():
    selected_item = contact_list.selection()
    if selected_item:
        item = contact_list.item(selected_item)
        name = item['values'][0]  # The first value is the contact's name
        confirm = messagebox.askyesno("Delete Contact", f"Are you sure you want to delete {name}?")
        if confirm:
            del address_book[name]  # Remove the contact from the address book
            save_contacts()  # Save changes
            view_contacts()  # Update the view
            clear_entries()
            messagebox.showinfo("Success", f"Contact {name} deleted successfully!")
    else:
        messagebox.showwarning("Select Contact", "Please select a contact to delete.")

# Search for contacts by any parameter
def search_contacts():
    search_query = entry_search.get().lower().strip()  # Trim spaces and convert to lowercase
    contact_list.delete(*contact_list.get_children())
    found = False
    for name, details in address_book.items():
        # Check if the search query matches any field: name, phone, email, or location
        if (search_query in name.lower() or
            search_query in details['Phone'].lower() or
            search_query in details['Email'].lower() or
            search_query in details['Location'].lower()):
            contact_list.insert("", "end", values=(name, details['Phone'], details['Email'], details['Location']))
            found = True
    if not found:
        messagebox.showinfo("Search Result", "No contacts found.")
    elif not search_query:
        messagebox.showinfo("Search Error", "Please enter a search query.")

# View contacts
def view_contacts():
    contact_list.delete(*contact_list.get_children())
    if address_book:
        for name, details in address_book.items():
            contact_list.insert("", "end", values=(name, details['Phone'], details['Email'], details['Location']))
    else:
        messagebox.showinfo("Info", "The address book is empty.")

# Sorting functionality for each column
def treeview_sort_column(tv, col, reverse):
    # Get the data in the selected column for all rows
    data = [(tv.set(child, col), child) for child in tv.get_children('')]
    # Sort data by the column clicked
    data.sort(reverse=reverse)

    # Rearrange items in sorted positions
    for index, (val, item) in enumerate(data):
        tv.move(item, '', index)

    # Reverse sort for the next click on this column
    tv.heading(col, command=lambda: treeview_sort_column(tv, col, not reverse))

# Populate the entry fields with the selected contact
def on_contact_select(event):
    selected_item = contact_list.selection()
    if selected_item:
        item = contact_list.item(selected_item)
        name, phone, email, location = item['values']
        entry_name.delete(0, tk.END)
        entry_name.insert(0, name)
        entry_phone.delete(0, tk.END)
        entry_phone.insert(0, phone)
        entry_email.delete(0, tk.END)
        entry_email.insert(0, email)
        entry_location.delete(0, tk.END)
        entry_location.insert(0, location)

# Clear input entries
def clear_entries():
    entry_name.delete(0, tk.END)
    entry_phone.delete(0, tk.END)
    entry_email.delete(0, tk.END)
    entry_location.delete(0, tk.END)

# Function to add a company logo at the top of the window
def add_logo(window, logo_path):
    try:
        logo = tk.PhotoImage(file=logo_path)
        label_logo = tk.Label(window, image=logo, bg="#f0f0f0")
        label_logo.image = logo
        label_logo.pack(pady=10)
    except Exception as e:
        messagebox.showerror("Error", f"Unable to load logo: {e}")

# Function to set background color for the entire window and components
def set_background_color(color_code="#A020F0"):
    root.configure(bg=color_code)  # Set background color for the main window
    frame_input.configure(bg=color_code)  # Set for input frame
    frame_search.configure(bg=color_code)  # Set for search frame
    frame_buttons.configure(bg=color_code)  # Set for buttons frame

    # Update background color of all labels
    label_name.configure(bg=color_code)
    label_phone.configure(bg=color_code)
    label_email.configure(bg=color_code)
    label_location.configure(bg=color_code)
    label_search.configure(bg=color_code)

# Create the main window
root = tk.Tk()
root.title("Address Book")
root.geometry("900x700")  # Adjusted window size for centering
root.configure(bg="#f0f0f0")

# Global variable to hold the contacts
address_book = load_contacts()

add_logo(root, "C:\\Users\\NickPetraglia\\Documents\\logo.png")

# Frame for input fields
frame_input = tk.Frame(root, bg="#f0f0f0", padx=10, pady=10)
frame_input.pack(pady=10)

# Configure grid to center the input forms
frame_input.grid_columnconfigure(0, weight=1)
frame_input.grid_columnconfigure(2, weight=1)

# Name Label and Entry
label_name = tk.Label(frame_input, text="Name:", font=('Arial', 12), bg="#f0f0f0")
label_name.grid(row=0, column=1, sticky="e", pady=5, padx=5)
entry_name = tk.Entry(frame_input, width=40, font=('Arial', 12))
entry_name.grid(row=0, column=2, pady=5, padx=5)

# Phone Label and Entry
label_phone = tk.Label(frame_input, text="Phone:", font=('Arial', 12), bg="#f0f0f0")
label_phone.grid(row=1, column=1, sticky="e", pady=5, padx=5)
entry_phone = tk.Entry(frame_input, width=40, font=('Arial', 12))
entry_phone.grid(row=1, column=2, pady=5, padx=5)

# Email Label and Entry
label_email = tk.Label(frame_input, text="Email:", font=('Arial', 12), bg="#f0f0f0")
label_email.grid(row=2, column=1, sticky="e", pady=5, padx=5)
entry_email = tk.Entry(frame_input, width=40, font=('Arial', 12))
entry_email.grid(row=2, column=2, pady=5, padx=5)

# Location Label and Entry
label_location = tk.Label(frame_input, text="Location:", font=('Arial', 12), bg="#f0f0f0")
label_location.grid(row=3, column=1, sticky="e", pady=5, padx=5)
entry_location = tk.Entry(frame_input, width=40, font=('Arial', 12))
entry_location.grid(row=3, column=2, pady=5, padx=5)

# Search Frame
frame_search = tk.Frame(root, bg="#f0f0f0", padx=10, pady=10)
frame_search.pack(pady=10)

# Search Label and Entry
label_search = tk.Label(frame_search, text="Search:", font=('Arial', 12), bg="#f0f0f0")
label_search.grid(row=0, column=1, sticky="e", pady=5, padx=5)
entry_search = tk.Entry(frame_search, width=30, font=('Arial', 12))
entry_search.grid(row=0, column=2, pady=5, padx=5)

# Search Button
btn_search = tk.Button(frame_search, text="Search", command=search_contacts, font=('Arial', 12), bg="#FFC107", fg="white", padx=20)
btn_search.grid(row=0, column=3, pady=5, padx=10)

# Buttons Frame
frame_buttons = tk.Frame(root, bg="#f0f0f0", padx=10, pady=10)
frame_buttons.pack(pady=10)

# Add/Update Contact Button
btn_add_update = tk.Button(frame_buttons, text="Add/Update Contact", command=add_or_update_contact, font=('Arial', 12), bg="#4CAF50", fg="white", padx=20)
btn_add_update.grid(row=0, column=1, padx=10)

# View Contacts Button
btn_view_contacts = tk.Button(frame_buttons, text="View Contacts", command=view_contacts, font=('Arial', 12), bg="#2196F3", fg="white", padx=20)
btn_view_contacts.grid(row=0, column=2, padx=10)

# Delete Contact Button
btn_delete = tk.Button(frame_buttons, text="Delete Contact", command=delete_contact, font=('Arial', 12), bg="#f44336", fg="white", padx=20)
btn_delete.grid(row=0, column=3, padx=10)

# Clear Button
btn_clear = tk.Button(frame_buttons, text="Clear Fields", command=clear_entries, font=('Arial', 12), bg="#A020F0", fg="white", padx=20)
btn_clear.grid(row=0, column=4, padx=10)

# Center the buttons by adjusting the grid layout
frame_buttons.grid_columnconfigure(0, weight=1)
frame_buttons.grid_columnconfigure(1, weight=1)
frame_buttons.grid_columnconfigure(2, weight=1)
frame_buttons.grid_columnconfigure(3, weight=1)
frame_buttons.grid_columnconfigure(4, weight=1)

# Treeview to display contacts
frame_contacts = tk.Frame(root, bg="#f0f0f0", padx=10, pady=10)
frame_contacts.pack(pady=10, fill="both", expand=True)

columns = ("Name", "Phone", "Email", "Location")
contact_list = ttk.Treeview(frame_contacts, columns=columns, show="headings", height=10)

# Define column headings and bind them to the sorting function
for col in columns:
    contact_list.heading(col, text=col, command=lambda _col=col: treeview_sort_column(contact_list, _col, False))

contact_list.column("Name", width=100)
contact_list.column("Phone", width=100)
contact_list.column("Email", width=180)
contact_list.column("Location", width=120)
contact_list.pack(padx=10, pady=10, fill="both", expand=True)

# Bind the contact selection event to populate fields
contact_list.bind('<<TreeviewSelect>>', on_contact_select)

# Apply styles for Treeview
style = ttk.Style()
style.configure("Treeview", font=('Arial', 12), rowheight=20)

# Run the main
view_contacts()  # Show the contacts when the GUI starts
root.mainloop()

