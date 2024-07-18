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


def show_message(title, message, type='info', master=None, custom=False):
    """
    Displays a message box of specified type. Optionally uses a custom dialog.

    Args:
        title (str): The title of the message box.
        message (str): The message to display.
        type (str): The type of message box ('info' or 'error').
        master (tk.Widget, optional): The parent widget if using a custom dialog.
        custom (bool): Whether to use the custom dialog.
    """
    if custom and master is not None:
        show_custom_message(master, title, message, type)
    elif type == 'info':
        tk.messagebox.showinfo(title, message)
    elif type == 'error':
        tk.messagebox.showerror(title, message)

def show_custom_message(master, title, message, type='info'):
    """
    Displays a custom message box with selectable text that adapts its height based on content length.
    Scrolls appear after 20 lines.

    Args:
        master (tk.Widget): The parent widget.
        title (str): The title of the message box.
        message (str): The message to display.
        type (str): The type of message box ('info' or 'error').
    """
    top = tk.Toplevel(master)
    top.title(title)

    screen_width = master.winfo_screenwidth()
    screen_height = master.winfo_screenheight()
    window_width = min(350, max(300, len(message) * 7))

    characters_per_line = window_width // 7
    estimated_lines = len(message) // characters_per_line + (1 if len(message) % characters_per_line else 0)
    window_height = min(150 + 18 * min(20, estimated_lines), screen_height - 100)

    x_coordinate = (screen_width - window_width) // 2
    y_coordinate = (screen_height - window_height) // 2
    top.geometry(f"{window_width}x{window_height}+{x_coordinate}+{y_coordinate}")

    top.transient(master)
    top.grab_set()

    bg_color = '#dddddd' if type == 'info' else '#ffdddd'

    text_frame = tk.Frame(top)
    text_frame.pack(fill='both', expand=True)
    text_scroll = tk.Scrollbar(text_frame)
    text_scroll.pack(side='right', fill='y')
    text_widget = tk.Text(text_frame, wrap='word', yscrollcommand=text_scroll.set,
                          background=bg_color, borderwidth=0, highlightthickness=0)
    text_widget.insert('end', message)
    text_widget.config(state='disabled', height=min(25, estimated_lines))
    text_widget.pack(pady=20, padx=20, fill='both', expand=True)
    text_scroll.config(command=text_widget.yview)

    def on_click(event):
        if event.widget is not text_widget:
            top.destroy()

    top.bind("<FocusOut>", lambda event: top.destroy())
    top.bind("<Button-1>", on_click)
    top.focus_set()
    top.wait_window()
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

