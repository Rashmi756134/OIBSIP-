import tkinter as tk
import string
import random

def generate_password():
    length = int(length_var.get())
    complexity = complexity_var.get()
    characters = string.ascii_letters  # Default to letters
    if complexity == "Medium":
        characters += string.digits
    elif complexity == "High":
        characters += string.punctuation
    password = ''.join(random.choice(characters) for _ in range(length))
    password_var.set(password)

root = tk.Tk()
root.geometry("370x310")
root.title("Rashmi's Random Password Generator")

length_label = tk.Label(root, text="Password Length:")
length_label.pack(pady=5)
length_var = tk.StringVar()
length_entry = tk.Entry(root, textvariable=length_var)
length_entry.pack(pady=5)

complexity_label = tk.Label(root, text="Complexity:")
complexity_label.pack(pady=5)
complexity_options = ["Low", "Medium", "High"]
complexity_var = tk.StringVar(root, "Low")  # Default to Low
complexity_menu = tk.OptionMenu(root, complexity_var, *complexity_options)
complexity_menu.pack(pady=5)

generate_button = tk.Button(root, text="Generate Password", command=generate_password)
generate_button.pack(pady=10)

password_var = tk.StringVar()
password_label = tk.Label(root, text="Generated Password:")
password_label.pack(pady=5)
password_entry = tk.Entry(root, textvariable=password_var, state="readonly", width=30)
password_entry.pack(pady=5)

root.mainloop()