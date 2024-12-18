import os
import pickle
import tkinter as tk
import keyboard
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
import ctypes

# Windows API Constants
GWL_EXSTYLE = -20
WS_EX_LAYERED = 0x00080000
WS_EX_TRANSPARENT = 0x00000020
LWA_ALPHA = 0x00000002

PERSISTENCE_FILE = "sticky_note_settings.pkl"  # File to store app settings


def make_window_click_through(hwnd):
    """Make the window click-through."""
    extended_style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
    ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, extended_style | WS_EX_LAYERED | WS_EX_TRANSPARENT)


def remove_click_through(hwnd):
    """Remove the click-through property from the window."""
    extended_style = ctypes.windll.user32.GetWindowLongW(hwnd, GWL_EXSTYLE)
    ctypes.windll.user32.SetWindowLongW(hwnd, GWL_EXSTYLE, extended_style & ~WS_EX_TRANSPARENT)


class StickyNoteApp:
    def __init__(self):
        # Load last saved settings
        self.settings = self.load_settings()

        self.window = tk.Tk()
        self.window.geometry("400x300")
        self.window.configure(bg=self.settings.get("last_color", "white"))  # Load saved color
        self.window.title("Sticky Note")  # Native title bar
        self.window.resizable(True, True)

        # Add the text area
        self.text = tk.Text(self.window, borderwidth=0, highlightthickness=0, bg=self.settings.get("last_color", "white"))
        self.text.pack(fill=tk.BOTH, expand=True)

        # Add the menu button at the bottom-right corner
        self.menu_button = tk.Menubutton(self.window, text="☰", relief=tk.FLAT, bg=self.settings.get("last_color", "white"),
                                         font=("Arial", 8), activebackground="lightgray")
        self.menu_button.place(relx=1.0, rely=1.0, anchor="se", x=-10, y=-10)

        self.options_menu = tk.Menu(self.menu_button, tearoff=0)
        self.click_through_var = tk.BooleanVar(value=False)
        self.options_menu.add_checkbutton(
            label="Click Through",
            variable=self.click_through_var,
            command=self.toggle_click_through
        )

        # Add color change feature
        color_menu = tk.Menu(self.options_menu, tearoff=0)
        colors = ["white", "lightgray", "lavender", "lightblue", "lightyellow", "lightgreen"]
        for index, color in enumerate(colors, start=1):
            color_menu.add_command(
                label=f"{index}. {color.capitalize()}",
                command=lambda c=color: self.change_app_color(c)
            )
        self.options_menu.add_cascade(label="Change Color", menu=color_menu)

        self.menu_button.config(menu=self.options_menu)

        self.local_folder = "Sticky Notes"
        if not os.path.exists(self.local_folder):
            os.makedirs(self.local_folder)

        self.is_transparent = False

        self.text.bind("<FocusIn>", self.on_focus_in)
        self.text.bind("<FocusOut>", self.on_focus_out)

        # Authenticate with Google Drive API
        self.drive_service = self.authenticate_google_drive()

        # Bind hotkeys
        keyboard.add_hotkey('ctrl+s', self.save_note)
        keyboard.add_hotkey('ctrl+alt+n', self.restore_focus)

        # Ensure the window is always on top
        self.window.attributes("-topmost", 1)

        self.window.mainloop()

    def change_app_color(self, color):
        """Change the entire app's color and save the setting."""
        self.window.configure(bg=color)
        self.text.configure(bg=color)
        self.menu_button.configure(bg=color)

        # Save the new color to settings
        self.settings["last_color"] = color
        self.save_settings()

    def toggle_click_through(self):
        """Toggle the click-through feature."""
        hwnd = ctypes.windll.user32.GetParent(self.window.winfo_id())
        if self.click_through_var.get():
            self.window.attributes("-alpha", 0.5)
            make_window_click_through(hwnd)
        else:
            self.window.attributes("-alpha", 0.5)
            remove_click_through(hwnd)
            self.is_transparent = False

    def authenticate_google_drive(self):
        """Authenticate with Google Drive API."""
        SCOPES = ['https://www.googleapis.com/auth/drive.file']
        creds = None

        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'credentials.json', SCOPES)
                creds = flow.run_local_server(port=0)

            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        return build('drive', 'v3', credentials=creds)

    def save_note_locally(self):
        """Save the note locally."""
        existing_files = os.listdir(self.local_folder)
        max_index = 0
        for file in existing_files:
            if file.startswith("note_cache_") and file.endswith(".txt"):
                index = int(file.split("_")[-1].split(".")[0])
                max_index = max(max_index, index)
        new_filename = f"note_cache_{max_index + 1}.txt"

        note_content = self.text.get("1.0", tk.END).strip()
        file_path = os.path.join(self.local_folder, new_filename)
        with open(file_path, "w") as file:
            file.write(note_content)

        print(f"Note saved locally as {new_filename}.")
        return new_filename

    def save_note_to_drive(self, local_filename):
        """Upload the note to Google Drive."""
        folder_name = "Sticky Notes"
        folder_id = None

        response = self.drive_service.files().list(
            q=f"name='{folder_name}' and mimeType='application/vnd.google-apps.folder' and trashed=false",
            spaces='drive'
        ).execute()

        if response.get('files'):
            folder_id = response['files'][0]['id']
        else:
            file_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            folder = self.drive_service.files().create(body=file_metadata, fields='id').execute()
            folder_id = folder.get('id')

        drive_filename = f"Sticky Note {int(local_filename.split('_')[-1].split('.')[0])}"
        file_metadata = {
            'name': drive_filename,
            'parents': [folder_id]
        }
        media = MediaFileUpload(os.path.join(self.local_folder, local_filename), mimetype='text/plain')
        self.drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()

        print(f"Note uploaded to Google Drive as {drive_filename}.")

    def save_note(self):
        """Save the note both locally and to Google Drive."""
        if self.window.focus_get():
            local_filename = self.save_note_locally()
            self.save_note_to_drive(local_filename)

    def restore_focus(self):
        """Restore the sticky note to full opacity and interactivity."""
        hwnd = ctypes.windll.user32.GetParent(self.window.winfo_id())
        remove_click_through(hwnd)
        self.window.attributes("-alpha", 1)
        self.is_transparent = False

    def on_focus_in(self, event):
        """Handle focus in event."""
        self.window.attributes("-alpha", 1)
        self.window.attributes("-topmost", 1)
        hwnd = ctypes.windll.user32.GetParent(self.window.winfo_id())
        remove_click_through(hwnd)
        self.is_transparent = False

    def on_focus_out(self, event):
        """Handle focus out event."""
        if not self.is_transparent:
            self.window.attributes("-alpha", 0.5)
            self.window.attributes("-topmost", 1)
            if self.click_through_var.get():
                hwnd = ctypes.windll.user32.GetParent(self.window.winfo_id())
                make_window_click_through(hwnd)

    def save_settings(self):
        """Save settings to a file."""
        with open(PERSISTENCE_FILE, "wb") as file:
            pickle.dump(self.settings, file)

    def load_settings(self):
        """Load settings from a file."""
        if os.path.exists(PERSISTENCE_FILE):
            with open(PERSISTENCE_FILE, "rb") as file:
                return pickle.load(file)
        return {}

    def on_close(self):
        """Handle app close event."""
        self.window.destroy()


if __name__ == "__main__":
    app = StickyNoteApp()
