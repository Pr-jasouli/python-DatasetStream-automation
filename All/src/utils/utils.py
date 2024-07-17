import tkinter as tk
from tkinter import ttk, filedialog, messagebox


def center_window(window, master, width=420, height=650):
    """
    Centers the window on the screen based on the master window's size.

    Args:
        window (tk.Toplevel): The window to be centered.
        master (tk.Tk): The master window.
        width (int): The width of the window to be centered.
        height (int): The height of the window to be centered.
    """
    x = master.winfo_x() + (master.winfo_width() - width) // 2
    y = master.winfo_y() + (master.winfo_height() - height) // 2
    window.geometry(f"{width}x{height}+{x}+{y}")


def create_styled_button(parent, text, command=None):
    """
    Creates a styled button with a specific text and command.

    Args:
        parent (tk.Widget): The parent widget for the button.
        text (str): The text to display on the button.
        command (callable, optional): The function to call when the button is clicked.

    Returns:
        ttk.Button: The created styled button.
    """
    return ttk.Button(parent, text=text, command=command, style='TButton')


def create_menu(window, menu_items):
    """
    Creates a menu in the window with given menu items.

    Args:
        window (tk.Tk): The window to attach the menu to.
        menu_items (list): A list of dictionaries, each containing 'label' and 'command' keys.
    """
    menubar = tk.Menu(window)
    window.config(menu=menubar)
    for item in menu_items:
        submenu = item.get("submenu")
        if submenu:
            cascade_menu = tk.Menu(menubar, tearoff=0, background='SystemButtonFace', fg='black')
            for subitem in submenu:
                cascade_menu.add_command(label=subitem["label"], command=subitem["command"])
            menubar.add_cascade(label=item["label"], menu=cascade_menu)
        else:
            menubar.add_command(label=item["label"], command=item["command"])


def select_file(callback, filetypes):
    """
    Handles selecting a file and passes the selected file path to the callback function.

    Args:
        callback (callable): The function to call with the selected file path.
        filetypes (list): A list of tuples specifying the allowed file types.
    """
    file_selected = filedialog.askopenfilename(filetypes=filetypes)
    if file_selected:
        callback(file_selected)


def show_message(title, message, type='info'):
    """
    Displays a message box of specified type.

    Args:
        title (str): The title of the message box.
        message (str): The message to display.
        type (str): The type of message box ('info' or 'error').
    """
    if type == 'info':
        messagebox.showinfo(title, message)
    elif type == 'error':
        messagebox.showerror(title, message)


def select_directory(entry_field):
    """
    Opens a directory dialog to select a folder and updates the entry field.

    Args:
        entry_field (ttk.Entry): The entry field to update with the selected directory path.
    """
    folder_selected = filedialog.askdirectory()
    if folder_selected:
        entry_field.delete(0, tk.END)
        entry_field.insert(0, folder_selected)

def clean_file_path(file_path):
    """
    Cleans up the file path by removing leading and trailing quotes.

    Args:
        file_path (str): The file path to clean.

    Returns:
        str: The cleaned file path.
    """
    return file_path.strip().strip('"')
