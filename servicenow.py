from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import TimeoutException, WebDriverException
import time
import os
import logging
import sys
import webbrowser
import tkinter as tk
from tkinter import ttk, messagebox
import platform
import requests
import zipfile
import tempfile
import re

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def create_properly_sized_dialog(title, width=600, height=400):
    """Create a dialog that's properly sized and positioned"""
    dialog = tk.Toplevel()
    dialog.title(title)
    
    # Set initial geometry
    dialog.geometry(f"{width}x{height}")
    
    # Center the dialog on the screen
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    x = (screen_width - width) // 2
    y = (screen_height - height) // 2
    dialog.geometry(f"{width}x{height}+{x}+{y}")
    
    # Make the dialog resizable
    dialog.resizable(True, True)
    
    # Set minimum size to ensure buttons are visible
    dialog.minsize(width, height)
    
    dialog.lift()  # Bring window to front
    dialog.focus_force()  # Force focus
    
    # Update method to fit content to window
    def fit_to_content():
        dialog.update_idletasks()  # Ensure all widgets are properly sized
        dialog.geometry("")  # Reset geometry to fit content
    
    dialog.fit_to_content = fit_to_content
    
    return dialog

def add_context_menu(widget):
    """Add context menu with copy/paste to a widget"""
    menu = tk.Menu(widget, tearoff=0)
    
    def popup(event):
        try:
            menu.tk_popup(event.x_root, event.y_root, 0)
        finally:
            menu.grab_release()
    
    menu.add_command(label="Cut", command=lambda: widget.event_generate("<<Cut>>"))
    menu.add_command(label="Copy", command=lambda: widget.event_generate("<<Copy>>"))
    menu.add_command(label="Paste", command=lambda: widget.event_generate("<<Paste>>"))
    
    widget.bind("<Button-3>", popup)  # Right-click
    
    return menu
def ensure_edge_driver():
    """Ensures Edge WebDriver is available and returns its path"""
    # Define driver filename based on platform
    if platform.system() == "Windows":
        driver_filename = "msedgedriver.exe"
    else:
        driver_filename = "msedgedriver"
    
    # Check if driver already exists in current directory
    current_dir = os.path.dirname(os.path.abspath(__file__))
    driver_path = os.path.join(current_dir, driver_filename)
    
    if os.path.exists(driver_path):
        logger.info(f"Edge WebDriver found at: {driver_path}")
        return driver_path
    
    # If we're running from a PyInstaller package
    if getattr(sys, 'frozen', False):
        # Look in the _MEIPASS directory
        if hasattr(sys, '_MEIPASS'):
            driver_path = os.path.join(sys._MEIPASS, driver_filename)
            if os.path.exists(driver_path):
                logger.info(f"Edge WebDriver found in PyInstaller bundle at: {driver_path}")
                return driver_path
    
    # If we get here, we need to download the driver
    try:
        # Get Edge version to download matching driver
        edge_version = get_edge_version()
        if not edge_version:
            logger.warning("Could not determine Edge version. Will use manual entry.")
            return None
            
        driver_path = download_edge_driver(edge_version, driver_path)
        return driver_path
    except Exception as e:
        logger.error(f"Failed to download Edge WebDriver: {str(e)}")
        return None

def get_edge_version():
    """Get the installed Edge browser version"""
    try:
        if platform.system() == "Windows":
            # Check registry for Edge version
            try:
                import winreg
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Edge\BLBeacon")
                version, _ = winreg.QueryValueEx(key, "version")
                logger.info(f"Edge version from registry: {version}")
                return version
            except:
                # Try PowerShell approach
                import subprocess
                cmd = r'powershell -command "Get-AppxPackage -Name *edge* | Select-Object -ExpandProperty Version"'
                proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
                stdout, stderr = proc.communicate()
                if stdout:
                    version = stdout.decode('utf-8').strip()
                    logger.info(f"Edge version from PowerShell: {version}")
                    return version
                
                # Try direct executable check
                edge_paths = [
                    r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe",
                    r"C:\Program Files\Microsoft\Edge\Application\msedge.exe"
                ]
                
                for path in edge_paths:
                    if os.path.exists(path):
                        app_dir = os.path.dirname(path)
                        version_file = os.path.join(app_dir, "version")
                        if os.path.exists(version_file):
                            with open(version_file, 'r') as f:
                                version = f.read().strip()
                                logger.info(f"Edge version from version file: {version}")
                                return version
                        
        elif platform.system() == "Darwin":  # macOS
            cmd = r'/Applications/Microsoft\ Edge.app/Contents/MacOS/Microsoft\ Edge --version'
            proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
            stdout, stderr = proc.communicate()
            if stdout:
                match = re.search(r'Microsoft Edge ([\d\.]+)', stdout.decode('utf-8'))
                if match:
                    version = match.group(1)
                    logger.info(f"Edge version: {version}")
                    return version
                    
        elif platform.system() == "Linux":
            cmd = 'microsoft-edge --version'
            proc = subprocess.Popen(cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE, shell=True)
            stdout, stderr = proc.communicate()
            if stdout:
                match = re.search(r'Microsoft Edge ([\d\.]+)', stdout.decode('utf-8'))
                if match:
                    version = match.group(1)
                    logger.info(f"Edge version: {version}")
                    return version
                    
    except Exception as e:
        logger.error(f"Error getting Edge version: {str(e)}")
        
    # If we couldn't get the version, return None
    logger.warning("Could not determine Edge version")
    return None
    
def download_edge_driver(edge_version, driver_path):
    """Download the appropriate EdgeDriver based on Edge version"""
    # Get major version
    major_version = edge_version.split('.')[0]
    
    # Determine system architecture and platform
    if platform.system() == "Windows":
        if platform.machine().endswith('64'):
            arch = "win64"
        else:
            arch = "win32"
    elif platform.system() == "Darwin":  # macOS
        arch = "mac64"
    else:  # Linux
        arch = "linux64"
    
    # URL for EdgeDriver downloads
    driver_url = f"https://msedgedriver.azureedge.net/{edge_version}/edgedriver_{arch}.zip"
    
    logger.info(f"Downloading EdgeDriver from: {driver_url}")
    
    try:
        # Create a temporary directory
        with tempfile.TemporaryDirectory() as tmp_dir:
            # Download the zip file
            zip_path = os.path.join(tmp_dir, "edgedriver.zip")
            
            response = requests.get(driver_url, stream=True)
            response.raise_for_status()  # Raise an exception for HTTP errors
            
            with open(zip_path, "wb") as f:
                for chunk in response.iter_content(chunk_size=8192):
                    f.write(chunk)
            
            # Extract the driver
            with zipfile.ZipFile(zip_path, 'r') as zip_ref:
                zip_ref.extractall(tmp_dir)
            
            # Find the driver in the extracted files
            if platform.system() == "Windows":
                extracted_driver = os.path.join(tmp_dir, "msedgedriver.exe")
            else:
                extracted_driver = os.path.join(tmp_dir, "msedgedriver")
            
            # Copy to desired location
            import shutil
            shutil.copy2(extracted_driver, driver_path)
            
            # Make executable on non-Windows
            if platform.system() != "Windows":
                os.chmod(driver_path, 0o755)
            
            logger.info(f"EdgeDriver downloaded and saved to: {driver_path}")
            return driver_path
            
    except Exception as e:
        logger.error(f"Failed to download EdgeDriver: {str(e)}")
        return None

def scrape_servicenow(identifier, is_asset=True):
    """Open browser, then provide option for manual entry"""
    # Get the logger instance at the function level
    logger = logging.getLogger(__name__)
    
    logger.info(f"Starting ServiceNow scrape for {'asset' if is_asset else 'serial'}: {identifier}")
    
    # Generate direct URL to the CMDB record
    if is_asset:
        # Direct URL to the specific asset
        url = f"https://globalfoundries.service-now.com/u_cmdb_ci_notebook.do?sysparm_query=asset_tag%3D{identifier}"
    else:
        # Direct URL to the specific serial number
        url = f"https://globalfoundries.service-now.com/u_cmdb_ci_notebook.do?sysparm_query=serial_number%3D{identifier}"
    
    # Create the dialog
    dialog = create_properly_sized_dialog("ServiceNow Asset Retrieval", 500, 400)
    
    # Instructions
    ttk.Label(dialog, text="Follow these steps to retrieve asset information:").pack(pady=10)
    
    step1_frame = ttk.LabelFrame(dialog, text="Step 1: Open and view the asset in your browser")
    step1_frame.pack(fill="x", padx=10, pady=10)
    
    ttk.Label(step1_frame, text="• Click 'Open CMDB' to launch your browser").pack(anchor="w", padx=20, pady=2)
    ttk.Label(step1_frame, text="• Log in to ServiceNow if necessary").pack(anchor="w", padx=20, pady=2)
    ttk.Label(step1_frame, text="• Ensure you can see the asset details page").pack(anchor="w", padx=20, pady=2)
    
    # Result storage
    result = [None]
    
    def open_browser():
        webbrowser.open(url)
        open_button.config(state="disabled")
        manual_button.config(state="normal")
        scrape_button.config(state="normal")
        status_label.config(text="Browser opened. Once you're logged in and can see the asset details, click 'Try Automatic Scraping'.")
        
    open_button = ttk.Button(step1_frame, text="Open CMDB", command=open_browser)
    open_button.pack(pady=10)
    
    step2_frame = ttk.LabelFrame(dialog, text="Step 2: Retrieve the asset information")
    step2_frame.pack(fill="x", padx=10, pady=10)
    
    ttk.Label(step2_frame, text="Option 1: Try to automatically scrape data from the open page").pack(anchor="w", padx=20, pady=2)
    ttk.Label(step2_frame, text="Option 2: Enter the data manually from the browser").pack(anchor="w", padx=20, pady=2)
    
    # Function for automatic scraping
    def attempt_scrape():
        nonlocal result
        
        scrape_button.config(state="disabled")
        status_label.config(text="Attempting to scrape... please wait")
        
        try:
            from selenium.webdriver.edge.service import Service
            
            # Get the driver path
            driver_path = ensure_edge_driver()
            
            # Configure Edge options
            options = webdriver.EdgeOptions()
            options.add_experimental_option('excludeSwitches', ['enable-logging'])
            options.add_argument("--no-sandbox")
            options.add_argument("--disable-dev-shm-usage")
            
            # Important: directly use the local port to avoid localhost resolution issues
            options.add_argument("--remote-debugging-port=0")
            
            try:
                # Create the driver with explicit service if we have a path
                if driver_path:
                    service = Service(driver_path)
                    driver = webdriver.Edge(service=service, options=options)
                else:
                    # Otherwise, let Selenium find the driver
                    driver = webdriver.Edge(options=options)
                
                try:
                    # Go to the URL directly
                    driver.get(url)
                    
                    # Wait for the page to load, looking for asset_tag field
                    wait = WebDriverWait(driver, 10)
                    
                    try:
                        # Look for the asset tag field - this tells us we're on the right page
                        asset_tag_field = wait.until(
                            EC.presence_of_element_located((By.ID, "u_cmdb_ci_notebook.asset_tag"))
                        )
                        asset_tag = asset_tag_field.get_attribute("value")
                        
                        # Initialize our data dictionary
                        asset_data = {
                            "asset_tag": asset_tag,
                            "cmdb_url": driver.current_url
                        }
                        
                        # Define field mappings 
                        field_mappings = {
                            "serial_number": "u_cmdb_ci_notebook.serial_number",
                            "hostname": "u_cmdb_ci_notebook.name",
                            "operational_status": "u_cmdb_ci_notebook.operational_status",
                            "install_status": "u_cmdb_ci_notebook.install_status",
                            "location": "u_cmdb_ci_notebook.location",
                            "ci_region": "u_cmdb_ci_notebook.u_region",
                            "owned_by": "u_cmdb_ci_notebook.owned_by",
                            "assigned_to": "sys_display.u_cmdb_ci_notebook.assigned_to",
                            "manufacturer": "u_cmdb_ci_notebook.manufacturer",
                            "model_id": "u_cmdb_ci_notebook.model_id",
                            "model_description": "sys_display.u_cmdb_ci_notebook.model_id",
                            "vendor": "u_cmdb_ci_notebook.vendor",
                            "warranty_expiration": "u_cmdb_ci_notebook.warranty_expiration",
                            "os": "u_cmdb_ci_notebook.os",
                            "os_version": "u_cmdb_ci_notebook.os_version",
                            "comments": "u_cmdb_ci_notebook.comments"
                        }
                        
                        # Extract all the fields we can find
                        for field_name, element_id in field_mappings.items():
                            try:
                                element = driver.find_element(By.ID, element_id)
                                value = ""
                                
                                # Different elements store their values differently
                                if element.tag_name == "input":
                                    value = element.get_attribute("value")
                                elif element.tag_name == "textarea":
                                    value = element.text
                                elif element.tag_name == "select":
                                    try:
                                        # For select elements, get the selected option text
                                        selected = element.find_element(By.CSS_SELECTOR, "option:checked")
                                        value = selected.text
                                    except:
                                        # Or just try to get the value
                                        value = element.get_attribute("value")
                                elif element.tag_name == "span":
                                    # For display elements like sys_display fields
                                    value = element.text
                                else:
                                    # Default case - try value attribute first, then text
                                    value = element.get_attribute("value") or element.text
                                
                                # Store the value
                                if value:
                                    asset_data[field_name] = value
                                    logger.info(f"Found {field_name}: {value}")
                            except Exception as field_error:
                                logger.warning(f"Could not find {field_name}: {str(field_error)}")
                        
                        # Success! Store the result
                        result[0] = asset_data
                        status_label.config(text="Scraping successful!")
                        dialog.after(1000, dialog.destroy)  # Close after 1 second
                        
                    except TimeoutException:
                        raise Exception("Could not find asset tag field - not on the correct page or not loaded")
                        
                finally:
                    driver.quit()
                    
            except WebDriverException as e:
                if "msedgedriver.exe" in str(e) and "executable" in str(e):
                    status_label.config(text="Edge driver not found. Use manual entry.")
                    logger.error(f"Edge driver error: {str(e)}")
                    manual_button.config(state="normal")
                else:
                    raise
                
        except Exception as e:
            logger.error(f"Scraping error: {str(e)}")
            status_label.config(text=f"Automatic scraping failed. Please use manual entry.")
            manual_button.config(state="normal")
    
    # Function to show manual entry
    def show_manual():
        dialog.destroy()  # Close the current dialog
        result[0] = show_manual_entry_form(identifier, is_asset, url)

     # Buttons
    button_frame = ttk.Frame(step2_frame)
    button_frame.pack(pady=10)
    
    scrape_button = ttk.Button(button_frame, text="Try Automatic Scraping", 
                              command=attempt_scrape, state="disabled")
    scrape_button.pack(side="left", padx=10)
    
    manual_button = ttk.Button(button_frame, text="Enter Details Manually", 
                              command=show_manual, state="disabled")
    manual_button.pack(side="left", padx=10)
    
    # Status label
    status_label = ttk.Label(dialog, text="")
    status_label.pack(pady=10)
    
    # Wait for dialog to close
    dialog.wait_window()
    
    # If no result was collected, show manual entry as fallback
    if result[0] is None:
        return show_manual_entry_form(identifier, is_asset, url)
    
    return result[0]

def show_manual_entry_form(identifier, is_asset, url):
    """Show form for manual data entry"""
    logger.info("Showing manual entry form")
    
    entry_window = create_properly_sized_dialog("Manual Asset Entry", 700, 700)
    
    # Instructions
    ttk.Label(entry_window, text="Please enter the asset information manually:").pack(pady=10)
    
    # Open browser button to reference the asset
    def open_browser():
        webbrowser.open(url)
    
    ttk.Button(entry_window, text="View in CMDB", command=open_browser).pack(pady=10)
    
    # Create scrollable frame
    main_frame = ttk.Frame(entry_window)
    main_frame.pack(fill="both", expand=True, padx=10, pady=10)
    
    canvas = tk.Canvas(main_frame)
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    
    scrollable_frame = ttk.Frame(canvas)
    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    
    # Add mouse wheel scrolling
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def _on_frame_enter(event):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    def _on_frame_leave(event):
        canvas.unbind_all("<MouseWheel>")
    
    main_frame.bind("<Enter>", _on_frame_enter)
    main_frame.bind("<Leave>", _on_frame_leave)
    
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    
    # Field entry widgets
    fields = [
        "asset_tag", "hostname", "serial_number", "operational_status",
        "install_status", "location", "ci_region", "owned_by", 
        "assigned_to", "manufacturer", "model_id", "model_description",
        "vendor", "warranty_expiration", "os", "os_version"
    ]
    
    field_labels = {
        "asset_tag": "Asset Tag (Required)",
        "hostname": "Hostname",
        "serial_number": "Serial Number",
        "operational_status": "Operational Status",
        "install_status": "Install Status",
        "location": "Location",
        "ci_region": "CI Region",
        "owned_by": "Owned By",
        "assigned_to": "Assigned To",
        "manufacturer": "Manufacturer",
        "model_id": "Model ID",
        "model_description": "Model Description",
        "vendor": "Vendor",
        "warranty_expiration": "Warranty Expiration",
        "os": "Operating System",
        "os_version": "OS Version"
    }
    
    entries = {}
    
    for i, field in enumerate(fields):
        ttk.Label(scrollable_frame, text=f"{field_labels[field]}:").grid(row=i, column=0, sticky="w", padx=10, pady=5)
        entry = ttk.Entry(scrollable_frame, width=40)
        entry.grid(row=i, column=1, padx=10, pady=5)
        entries[field] = entry
        
        # Add context menu for copy/paste
        add_context_menu(entry)
        
        # Pre-fill the identifier
        if field == "asset_tag" and is_asset:
            entry.insert(0, identifier)
        elif field == "serial_number" and not is_asset:
            entry.insert(0, identifier)
    
    # Comments field (multiline)
    ttk.Label(scrollable_frame, text="Comments:").grid(row=len(fields), column=0, sticky="nw", padx=10, pady=5)
    comments_text = tk.Text(scrollable_frame, width=40, height=5)
    comments_text.grid(row=len(fields), column=1, padx=10, pady=5)
    
    # Add context menu for comments
    add_context_menu(comments_text)
    
    # Result storage
    result = [None]
    
    # Submit function
    def submit_data():
        asset_data = {field: entries[field].get().strip() for field in fields}
        asset_data['comments'] = comments_text.get("1.0", "end-1c")
        asset_data['cmdb_url'] = url
        
        if not asset_data['asset_tag']:
            messagebox.showerror("Error", "Asset Tag is required")
            return
        
        result[0] = asset_data
        entry_window.destroy()
    
    ttk.Button(entry_window, text="Submit Asset Information", command=submit_data).pack(pady=20)
    
    # Wait for window to close
    entry_window.wait_window()
    
    return result[0]   
