import os
import sys
import tkinter as tk
import tkinter.ttk as ttk
from tkinter import simpledialog, messagebox
import json
import ctypes # Import ctypes
import logging
from ui.app import InventoryApp

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s') # Configure basic logging if not done elsewhere
logger = logging.getLogger(__name__) 

if sys.platform == 'win32':
    try:
        # Try setting process DPI awareness (Windows 8.1+)
        ctypes.windll.shcore.SetProcessDpiAwareness(2) # 2 = Per Monitor DPI Aware v2
        print("DPI Awareness set to Per Monitor v2.")
    except (AttributeError, OSError):
        try:
             # Fallback for older Windows versions (Vista+)
             ctypes.windll.user32.SetProcessDPIAware()
             print("DPI Awareness set using SetProcessDPIAware.")
        except (AttributeError, OSError):
             print("Warning: Could not set DPI awareness.")
             

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)

def get_config_file_path():
    """Get path to app_config.json"""
    if getattr(sys, 'frozen', False):
        # If the application is run as a bundle
        # Use sys.executable to find the persistent location
        application_path = os.path.dirname(sys.executable)
    else:
        # Running as a script
        try:
            application_path = os.path.dirname(os.path.abspath(__file__))
        except NameError:
             application_path = os.path.abspath(".")

    return os.path.join(application_path, "app_config.json")

def load_or_create_config():
    """Load existing config or create new one with defaults"""
    print("Starting load_or_create_config")
    config_path = get_config_file_path()
    print(f"Config path: {config_path}")

    # Default configuration
    default_config = {
        "site": None,  # Will be prompted on first run
        "version": "1.0.1", # Example version, update as needed
    }

    # Check if config file exists
    if os.path.exists(config_path):
        print("Config file exists, loading")
        try:
            with open(config_path, 'r') as f:
                config = json.load(f)
                # Update with any missing default values
                for key, value in default_config.items():
                    if key not in config:
                        config[key] = value
                print(f"Loaded config: {config}")
                return config
        except Exception as e:
            print(f"Error loading config: {e}")
            # If loading fails, fall back to defaults but keep existing site if possible
            existing_site = config.get("site", None) if 'config' in locals() else None
            default_config["site"] = existing_site
            return default_config
    else:
        # Create new config file with defaults
        print("Config file does not exist, creating with defaults")
        try:
            # Ensure the directory exists before writing
            os.makedirs(os.path.dirname(config_path), exist_ok=True)
            with open(config_path, 'w') as f:
                json.dump(default_config, f, indent=2)
            return default_config
        except Exception as e:
            print(f"Error creating config file: {e}")
            return default_config # Return defaults even if creation fails

def save_config(config):
    """Save configuration to file"""
    print(f"Saving config: {config}")
    config_path = get_config_file_path()
    try:
        # Ensure the directory exists before writing
        os.makedirs(os.path.dirname(config_path), exist_ok=True)
        with open(config_path, 'w') as f:
            json.dump(config, f, indent=2)
        return True
    except Exception as e:
        print(f"Error saving config: {e}")
        return False

def main():
    print("Starting main function")
    logger.info

    # No need to change working directory if using absolute paths derived from sys.executable

    # Create initial tkinter window
    print("Creating Tkinter root window")
    root = tk.Tk()
    # Make the window icon themeable (optional but good practice)
    try:
        # Use themed icon if available (replace 'app_icon.ico' with your icon file)
        # Note: Ensure icon file is included using PyInstaller's --add-data or --icon
        root.iconbitmap(resource_path('app_icon.ico'))
    except Exception:
         print("Icon not found or not supported, using default.")


    # Load configuration
    print("Loading configuration")
    config = load_or_create_config()

    # Check if site is configured
    if config["site"] is None:
        print("Site not configured, creating site selection dialog")

        # Configure the window for site selection
        root.title("IT Bench Inventory - Site Configuration")

        # Calculate window size and position
        window_width = 450
        window_height = 550  # Adjusted height

        # Get screen dimensions
        screen_width = root.winfo_screenwidth()
        screen_height = root.winfo_screenheight()

        # Calculate position to center the window
        position_x = int(screen_width/2 - window_width/2)
        position_y = int(screen_height/2 - window_height/2)

        # Set geometry, make resizable, and set minimum size
        root.geometry(f"{window_width}x{window_height}+{position_x}+{position_y}")
        root.resizable(True, True) # Make the site selection window resizable
        root.minsize(window_width, window_height) # Set minimum size

        # Create a container frame with padding, make it expand
        container = tk.Frame(root, padx=20, pady=20)
        container.pack(fill="both", expand=True) # Make container fill the window

        # Configure container's grid weights if using grid for layout within it
        # container.rowconfigure(0, weight=0) # Header
        # container.rowconfigure(1, weight=1) # Selection frame (allow expansion)
        # container.rowconfigure(2, weight=0) # Button frame
        # container.columnconfigure(0, weight=1)

        # Header section
        header_frame = tk.Frame(container)
        header_frame.pack(fill="x", pady=(0, 20)) # Use pack, simpler for linear layout

        tk.Label(header_frame, text="IT Bench Inventory Manager", font=("", 16, "bold")).pack()
        tk.Label(header_frame, text="First-time Setup", font=("", 12)).pack(pady=(5, 0))

        # Site selection section - allow this frame to expand vertically
        selection_frame = tk.Frame(container)
        selection_frame.pack(fill="both", expand=True) # Expand this frame

        tk.Label(selection_frame, text="Please select your site:", font=("", 12, "bold")).pack(anchor="w")
        tk.Label(selection_frame, text="This setting determines which bench inventory\nyou'll be managing.",
                font=("", 10), justify="left").pack(anchor="w", pady=(0, 10))

        # Sites list
        sites = ['AUS', 'BTV', 'EFK', 'SCG', 'MALTA']
        selected_site = tk.StringVar()
        selected_site.set(sites[0])  # Default to first site

        # Radio buttons in their own frame
        radio_frame = tk.Frame(selection_frame)
        radio_frame.pack(fill="x", pady=5) # Fill horizontally

        for site in sites:
            # Use ttk.Radiobutton for better theme support
            rb = ttk.Radiobutton(radio_frame, text=site, variable=selected_site,
                              value=site) # Removed font setting to use theme default
            rb.pack(anchor="w", padx=40, pady=3) # Adjust padding

        # Button frame at the bottom (doesn't need to expand)
        button_frame = tk.Frame(container)
        button_frame.pack(side="bottom", fill="x", pady=(10, 0)) # Add padding top

        def on_continue():
            site = selected_site.get()
            if site:
                config["site"] = site
                if save_config(config):
                    print(f"Site configured: {site}")

                    # Clear the window
                    for widget in root.winfo_children():
                        widget.destroy()

                    # Restore original window settings for main app
                    root.title("IT Bench Inventory Manager")
                    # Remove fixed geometry for the main app window
                    # root.geometry("800x600")
                    root.resizable(True, True) # Ensure main window is resizable
                    # Optionally set a minimum size for the main window
                    root.minsize(600, 400)

                    # Center the main window (optional)
                    root.update_idletasks() # Update geometry info
                    win_width = root.winfo_width()
                    win_height = root.winfo_height()
                    scr_width = root.winfo_screenwidth()
                    scr_height = root.winfo_screenheight()
                    pos_x = int(scr_width / 2 - win_width / 2)
                    pos_y = int(scr_height / 2 - win_height / 2)
                    root.geometry(f"+{pos_x}+{pos_y}")


                    # Initialize the app with the configuration
                    app = InventoryApp(root, config)
                else:
                     messagebox.showerror("Configuration Error",
                                   f"Failed to save site configuration to {get_config_file_path()}.\nPlease check file permissions.", parent=root)
            else:
                messagebox.showerror("Configuration Error",
                                   "Site selection is required to run the application.", parent=root)

        # Create a themed button
        # Use ttk.Button for better theme integration
        continue_button = ttk.Button(button_frame, text="Continue with Selected Site",
                                   command=on_continue,
                                   width=30) # Adjusted width slightly
        # Pack button centered within its frame
        continue_button.pack(pady=10)

    else:
        # Site is already configured, proceed normally
        print(f"Using configured site: {config['site']}")

        # Configure window for main application
        root.title("IT Bench Inventory Manager")
        root.resizable(True, True)
        root.minsize(700, 500) # Keep your minimum size

        # Center the main window
        root.update_idletasks() # Update geometry info
        min_w = root.winfo_reqwidth()
        min_h = root.winfo_reqheight()
        scr_width = root.winfo_screenwidth()
        scr_height = root.winfo_screenheight()

        # ---> MODIFY INITIAL HEIGHT CALCULATION HERE <---
        # Add a small buffer (e.g., 30 pixels) to the required height
        status_bar_height_buffer = 30
        start_w = max(min_w, 1000) # Your desired starting width
        # Ensure start_h is at least min_h + buffer, or your preferred start height (600)
        start_h = max(min_h + status_bar_height_buffer, 600)
        # ---> END MODIFICATION <---

        pos_x = int(scr_width / 2 - start_w / 2)
        pos_y = int(scr_height / 2 - start_h / 2)
        # Prevent negative coordinates on smaller screens
        pos_x = max(0, pos_x)
        pos_y = max(0, pos_y)
        root.geometry(f"{start_w}x{start_h}+{pos_x}+{pos_y}")

        # Initialize app with the configuration
        print("Initializing app with config")
        app = InventoryApp(root, config)

    def on_closing():
        """Handle window closing event."""
        logger.info("Application closing...")
        # Check if 'app' exists in the global scope before trying to access it
        global app  # Access the global app variable
        if 'app' in globals() and app is not None and hasattr(app, 'db'):
            app.db.close_db()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)
    print("Starting mainloop")
    root.mainloop()

if __name__ == "__main__":
    try:
        main()
    except Exception as e:
        # Log the exception with traceback
        import traceback
        import logging
        logging.basicConfig(filename='error.log', level=logging.ERROR,
                            format='%(asctime)s:%(levelname)s:%(message)s')
        logging.exception("Unhandled exception in main:") # Logs the traceback

        # Show error message to user
        try:
            # Try showing message box - might fail if Tkinter is broken
            messagebox.showerror("Fatal Error",
                               f"A critical error occurred: {str(e)}\n\n"
                               "Details have been logged to error.log.\n"
                               "The application will now close.")
        except Exception as tk_err:
             print(f"CRITICAL ERROR: {e}\nTraceback:\n{traceback.format_exc()}")
             print(f"Could not display Tkinter error message: {tk_err}")
        sys.exit(1) # Exit after fatal error