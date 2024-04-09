import tkinter as tk
from tkinter import filedialog, messagebox
import os
import pyperclip #type: ignore
from winreg import ConnectRegistry, OpenKey, QueryValueEx, HKEY_LOCAL_MACHINE
import traceback
from Registry import Registry #type: ignore

def decode_product_key(digital_product_id_bytes):
    digital_product_id_bytes = bytearray(digital_product_id_bytes)
    key_offset = 52
    chars = "BCDFGHJKMPQRTVWXY2346789"
    product_key = ""
    
    for i in range(24, -1, -1):
        cur = 0
        for j in range(14, -1, -1):
            cur *= 256
            cur += digital_product_id_bytes[j + key_offset]
            digital_product_id_bytes[j + key_offset] = cur // 24
            cur %= 24
        product_key = chars[cur] + product_key
        if i % 5 == 0 and i != 0:
            product_key = "-" + product_key
    return product_key

def get_local_digital_product_id():
    try:
        reg = ConnectRegistry(None, HKEY_LOCAL_MACHINE)
        key = OpenKey(reg, r"SOFTWARE\Microsoft\Windows NT\CurrentVersion")
        value, _ = QueryValueEx(key, "DigitalProductId")
        return value
    except Exception as e:
        messagebox.showerror("Error", f"Failed to access local registry: {str(e)}")
        return None

def show_product_key_window(product_key):
    def copy_to_clipboard():
        pyperclip.copy(product_key)
        messagebox.showinfo("Copied", "Product key copied to clipboard.")

    def save_to_file():
        file_path = filedialog.asksaveasfilename(defaultextension=".txt", filetypes=[("Text Files", "*.txt")])
        if file_path:
            with open(file_path, "w") as file:
                file.write(product_key)
            messagebox.showinfo("Saved", "Product key saved to file.")

    key_window = tk.Toplevel(root)
    key_window.title("Product Key")
    tk.Label(key_window, text="Product Key found. Product Key:", padx=10).pack(pady=(10,0))
    key_text = tk.Text(key_window, height=1, width=50)
    key_text.pack(padx=10, pady=5)
    key_text.insert(tk.END, product_key)
    key_text.configure(state='disabled')
    tk.Button(key_window, text="Copy to Clipboard", command=copy_to_clipboard).pack(side=tk.LEFT, padx=(20,10), pady=10)
    tk.Button(key_window, text="Save to File", command=save_to_file).pack(side=tk.RIGHT, padx=(10,20), pady=10)

def handle_local_key():
    digital_product_id_bytes = get_local_digital_product_id()
    if digital_product_id_bytes:
        product_key = decode_product_key(digital_product_id_bytes)
        show_product_key_window(product_key)

def setup_external_key_ui():
    for widget in root.winfo_children():
        widget.destroy()

    tk.Label(root, text="Enter the Windows folder path of the external computer/OS drive.", padx=10).pack(pady=(10,5))
    instruction_text = ("If you want to find the key for your current computer, "
                        "then go back and select the 'Get Product Key of This Computer'.")
    tk.Label(root, text=instruction_text, padx=10, wraplength=400).pack(pady=(10,5))
    path_frame = tk.Frame(root)
    path_frame.pack(pady=5, padx=10)
    
    global path_entry
    path_entry = tk.Entry(path_frame, width=50)
    path_entry.insert(0, "C:/Windows")
    path_entry.pack(side=tk.LEFT, fill=tk.X, expand=True)
    tk.Button(path_frame, text="Browse...", command=browse_folder).pack(side=tk.RIGHT, padx=(10,0))

    tk.Button(root, text="Find Product Key", command=handle_external_key, padx=10).pack(pady=5)
    tk.Button(root, text="Back", command=main_menu, padx=10).pack()

def browse_folder():
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        path_entry.delete(0, tk.END)
        path_entry.insert(0, folder_selected)

def handle_external_key():
    windows_folder = path_entry.get()
    hive_path = os.path.join(windows_folder, "System32", "config", "SOFTWARE")
    digital_product_id_bytes = get_digital_product_id(hive_path)
    if digital_product_id_bytes:
        product_key = decode_product_key(digital_product_id_bytes)
        show_product_key_window(product_key)

def get_digital_product_id(filepath):
    try:
        reg = Registry.Registry(filepath)
        key = reg.open("Microsoft\\Windows NT\\CurrentVersion")
        value = key.value("DigitalProductId").value()
        return value
    except PermissionError:
        messagebox.showerror(
            "Permission Denied",
            "Oops! We couldn't access the registry hive because of permission issues.\n\n"
            "Here's how you can try to fix it:\n"
            "1. Close this program.\n"
            "2. Right-click on the program icon and select 'Run as administrator'.\n"
            "3. Try selecting the Windows folder again.\n\n"
            "If you're trying to access a Windows folder from another computer or a backup:\n"
            "1. Right-click on the folder you're trying to access.\n"
            "2. Select 'Properties' and then go to the 'Security' tab.\n"
            "3. Click 'Edit...' to change permissions, then 'Add...' to add your user.\n"
            "4. Type your username into the box and click 'Check Names', then 'OK'.\n"
            "5. Make sure your user is selected, then check the box under 'Allow' next to 'Full control'.\n"
            "6. Click 'Apply', then 'OK' to close the windows.\n"
            "7. Try running this program again and select the folder.\n\n"
            "These steps might require administrative access to your computer."
        )
    except Exception as e:
        messagebox.showerror(
            "Error",
            f"An unexpected error occurred:\n{str(e)}\n\n{traceback.format_exc()}"
        )
    return None

def main_menu():
    for widget in root.winfo_children():
        widget.destroy()
    tk.Label(root, text="Prawn's Windows Product Key Finder", font=("Arial", 14)).pack(pady=(20,10))
    tk.Button(root, text="Get Product Key of This Computer", command=handle_local_key, width=50, pady=5).pack(pady=10)
    tk.Button(root, text="Get Product Key From Another Computer", command=setup_external_key_ui, width=50, pady=5).pack(pady=10)

root = tk.Tk()
root.title("Prawn's Windows Product Key Finder")
root.geometry("500x250")
root.resizable(False, False)
main_menu()
root.mainloop()
