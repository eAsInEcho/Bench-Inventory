import tkinter as tk
from tkinter import ttk, messagebox, Toplevel, Text, Listbox, Scrollbar, Menu, simpledialog, scrolledtext 
import os
import webbrowser
import logging
import threading
import re 
from ui.utils import add_context_menu, add_mousewheel_scrolling
from services.servicenow import scrape_servicenow 

logger = logging.getLogger(__name__)

def format_timestamp(timestamp):
    """Format timestamp to remove seconds and shorten timezone"""
    if timestamp is None or timestamp == 'Unknown' or timestamp == '':
        return ''

    timestamp_str = str(timestamp)

    try:
        # Simpler approach: find first occurrence of format like YYYY-MM-DD HH:MM
        match = re.search(r'(\d{4}-\d{2}-\d{2}\s+\d{2}:\d{2})', timestamp_str)
        if match:
            # Optionally add timezone if present and simple format like +HH or -HH
            tz_match = re.search(r'([-+]\d{2}):\d{2}', timestamp_str)
            if tz_match:
                return f"{match.group(1)} {tz_match.group(1)}"
            else:
                # Check for Z timezone
                if 'Z' in timestamp_str.upper():
                     return f"{match.group(1)} Z"
                # Check for simple offset like +0500
                tz_simple_match = re.search(r'([-+]\d{2})(\d{2})', timestamp_str)
                if tz_simple_match:
                     return f"{match.group(1)} {tz_simple_match.group(1)}"

            return match.group(1) # Return date/time only if no simple TZ found

        # Fallback if main pattern not found
        return timestamp_str
    except Exception as e:
        print(f"Error formatting timestamp {timestamp_str}: {str(e)}")
        return timestamp_str # Return original on error

def create_properly_sized_dialog(title, min_width=600, min_height=400, parent=None):
    """Create a dialog that's properly sized, positioned, resizable, and has a minimum size."""
    dialog = tk.Toplevel(parent) if parent else tk.Toplevel()
    dialog.title(title)

    # Remove initial geometry setting
    # dialog.geometry(f"{width}x{height}")

    # Center the dialog based on minimum size or desired initial size
    initial_width = min_width # Start at minimum
    initial_height = min_height
    screen_width = dialog.winfo_screenwidth()
    screen_height = dialog.winfo_screenheight()
    x = max(0, (screen_width - initial_width) // 2) # Ensure x, y are not negative
    y = max(0, (screen_height - initial_height) // 2)
    # Position the window initially
    dialog.geometry(f"+{x}+{y}")

    # Make the dialog resizable
    dialog.resizable(True, True)
    # Set MINIMUM size
    dialog.minsize(min_width, min_height)

    dialog.lift()
    dialog.focus_force()

    # Configure the dialog's main grid layout to allow content expansion
    dialog.columnconfigure(0, weight=1)
    dialog.rowconfigure(0, weight=1)  # Content area (row 0) will expand
    dialog.rowconfigure(1, weight=0)  # Button area (row 1) will stay fixed height

    return dialog

def show_check_in_out_dialog(db, asset_data, default_status="in", current_status=None, callback=None, site_config=None):
    """Show dialog to check asset in or out"""
    # Use create_properly_sized_dialog for consistent behavior
    # Adjusted minimum size
    check_window = create_properly_sized_dialog("Check In/Out Asset", min_width=450, min_height=350)

    # --- Asset ID Handling --- (Seems correct)
    asset_id = asset_data.get('asset_id', asset_data.get('asset_tag', None))
    if not asset_id:
        logger.error("No asset_id or asset_tag found in asset_data")
        messagebox.showerror("Error", "Cannot identify asset ID", parent=check_window)
        check_window.destroy()
        return

    logger.info(f"Showing check in/out dialog for asset: {asset_id}")

    # --- Status/Flag Fetching --- (Seems correct)
    if not current_status:
        current_status = db.get_asset_current_status(asset_id)
    current_state = current_status.get('status', 'unknown') if current_status else 'unknown'
    last_action_time = format_timestamp(current_status.get('timestamp', 'Unknown') if current_status else 'Unknown')
    
    # Get both flag statuses
    flag_status = db.get_flag_status(asset_id)
    is_flagged = flag_status and flag_status.get('flag_status', False)
    
    expiry_status = db.get_expiry_flag_status(asset_id)
    is_expiring = expiry_status and expiry_status.get('expiry_flag_status', False)

    # --- Main Layout ---
    # Create a main content frame that will fill the expandable area (row 0)
    content_frame = ttk.Frame(check_window, padding="10")
    content_frame.grid(row=0, column=0, sticky="nsew") # Expand content frame
    # Configure content_frame's layout if needed (e.g., if using grid inside)
    content_frame.columnconfigure(0, weight=1) # Make child widgets expand horizontally

    # --- Widgets within content_frame ---
    # Use pack for simple vertical layout within the content frame
    info_frame = ttk.LabelFrame(content_frame, text="Asset Information", padding="5")
    info_frame.pack(fill="x", padx=0, pady=5)
    ttk.Label(info_frame, text=f"Asset Tag: {asset_id}").pack(anchor="w", padx=5, pady=2)
    ttk.Label(info_frame, text=f"Model: {asset_data.get('manufacturer', '')} {asset_data.get('model_description', '')}").pack(anchor="w", padx=5, pady=2)
    ttk.Label(info_frame, text=f"S/N: {asset_data.get('serial_number', '')}").pack(anchor="w", padx=5, pady=2)
    
    # Add lease information if available
    lease_start = asset_data.get('lease_start_date')
    lease_end = asset_data.get('lease_maturity_date')
    
    if lease_start or lease_end:
        lease_frame = ttk.LabelFrame(content_frame, text="Lease Information", padding="5")
        lease_frame.pack(fill="x", padx=0, pady=5)
        
        if lease_start:
            ttk.Label(lease_frame, text=f"Lease Start: {lease_start}").pack(anchor="w", padx=5, pady=2)
        
        if lease_end:
            ttk.Label(lease_frame, text=f"Lease Maturity: {lease_end}").pack(anchor="w", padx=5, pady=2)
            
            # Calculate days remaining
            try:
                from datetime import datetime
                
                # Try different date formats
                maturity_date = None
                for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d']:
                    try:
                        maturity_date = datetime.strptime(lease_end, fmt).date()
                        break
                    except (ValueError, TypeError):
                        continue
                
                if maturity_date:
                    today = datetime.now().date()
                    days_remaining = (maturity_date - today).days
                    
                    status_text = ""
                    if days_remaining < 0:
                        status_text = f"EXPIRED ({abs(days_remaining)} days ago)"
                    elif days_remaining <= 90:
                        status_text = f"EXPIRING SOON ({days_remaining} days remaining)"
                    else:
                        status_text = f"Active ({days_remaining} days remaining)"
                    
                    ttk.Label(lease_frame, text=f"Lease Status: {status_text}").pack(anchor="w", padx=5, pady=2)
            except Exception as e:
                logger.error(f"Error calculating lease days: {e}")

    status_frame = ttk.LabelFrame(content_frame, text="Current Status", padding="5")
    status_frame.pack(fill="x", padx=0, pady=5)

    indicator_frame = ttk.Frame(status_frame)
    indicator_frame.pack(fill="x", padx=5, pady=5)

    indicator_size = 20
    indicator_canvas = tk.Canvas(indicator_frame, width=indicator_size, height=indicator_size, highlightthickness=0)
    indicator_canvas.pack(side="left", padx=(0, 5))
    indicator_color = "#4CAF50" if current_state == "in" else "#F44336"
    indicator_canvas.create_oval(2, 2, indicator_size-2, indicator_size-2, fill=indicator_color, outline="")

    site_text = current_status.get('site', '') if current_status else ''
    status_label_text = f"Status: {current_state.upper()}"
    if site_text and site_text != 'Out':
        status_label_text += f" at {site_text}"
    status_label_text += f" ({last_action_time})"
    status_label = ttk.Label(indicator_frame, text=status_label_text, font=("", 10, "bold"))
    status_label.pack(side="left", padx=5)

    # Show flag indicators if applicable
    flags_frame = ttk.Frame(status_frame)
    
    # Regular flag
    if is_flagged:
        flags_frame.pack(fill="x", padx=5, pady=(0,5)) # Pack below main status line
        flag_text = f"⚠️ FLAGGED by {flag_status.get('flag_tech', 'Unknown')}: {flag_status.get('flag_notes', '')[:30]}"
        flag_label = ttk.Label(flags_frame, text=flag_text,
                              foreground="#FF5722", font=("", 9, "italic"))
        flag_label.pack(anchor="w")
    
    # DaaS expiry flag
    if is_expiring:
        # If flags_frame is not yet packed, pack it
        if not is_flagged:
            flags_frame.pack(fill="x", padx=5, pady=(0,5))
        
        # Add a separator if both flags are present
        if is_flagged:
            ttk.Separator(flags_frame, orient="horizontal").pack(fill="x", pady=3)
        
        # Show expiry info
        days_left = ""
        if lease_end:
            try:
                from datetime import datetime
                
                # Try different date formats
                maturity_date = None
                for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d']:
                    try:
                        maturity_date = datetime.strptime(lease_end, fmt).date()
                        break
                    except (ValueError, TypeError):
                        continue
                
                if maturity_date:
                    today = datetime.now().date()
                    days_remaining = (maturity_date - today).days
                    days_left = f" ({days_remaining} days left)"
            except Exception as e:
                logger.error(f"Error calculating lease days: {e}")
        
        expiry_text = f"🕓 DaaS EXPIRING{days_left}"
        expiry_label = ttk.Label(flags_frame, text=expiry_text,
                               foreground="#9C27B0", font=("", 9, "italic"))
        expiry_label.pack(anchor="w")

    # --- Action Functions (check_in, check_out, view_details, view_cmdb, toggle_flag) ---
    # The logic within these seems okay, but ensure dialogs they open are also scalable.
    # Let's focus on the check-in/out technician dialog within this function:
    def record_status_change(status):
        if status == current_state:
            messagebox.showwarning("Invalid Action", f"Asset is already checked {status}.", parent=check_window)
            return

        # Technician prompt dialog
        # Use create_properly_sized_dialog
        tech_dialog = create_properly_sized_dialog(f"Check {status.upper()} - {asset_id}", min_width=400, min_height=300, parent=check_window)

        # Main content frame for this sub-dialog
        tech_content = ttk.Frame(tech_dialog, padding="10")
        tech_content.grid(row=0, column=0, sticky="nsew") # Expand content
        tech_dialog.rowconfigure(0, weight=1) # Allow content row to expand
        tech_dialog.columnconfigure(0, weight=1) # Allow content col to expand

        # Configure tech_content grid
        tech_content.rowconfigure(3, weight=1) # Make notes Text widget row expandable
        tech_content.columnconfigure(0, weight=1) # Make widgets fill horizontally

        ttk.Label(tech_content, text="Technician:").grid(row=0, column=0, sticky="w", padx=0, pady=(0,2))
        tech_entry = ttk.Entry(tech_content, width=30)
        tech_entry.grid(row=1, column=0, sticky="ew", padx=0, pady=(0,5)) # sticky='ew' to expand horizontally
        tech_entry.insert(0, os.getenv('USERNAME', ''))
        add_context_menu(tech_entry)

        # Determine site based on action (logic seems okay)
        site_value = site_config.get('site', 'Unknown') if site_config and status == "in" else "Out"
        site_display = site_value if status == "in" else f"Out (from {site_config.get('site', 'Unknown') if site_config else 'Unknown'})"
        ttk.Label(tech_content, text=f"Site: {site_display}").grid(row=2, column=0, sticky="w", padx=0, pady=5)

        ttk.Label(tech_content, text="Notes:").grid(row=3, column=0, sticky="nw", padx=0, pady=(5,2))
        # Use scrolledtext for notes
        notes_text = scrolledtext.ScrolledText(tech_content, width=40, height=5, wrap="word")
        notes_text.grid(row=4, column=0, sticky="nsew", padx=0, pady=(0,5)) # Expand notes text area
        add_context_menu(notes_text)
        notes_text.focus() # Focus notes

        # --- Submit/Cancel Buttons for tech_dialog (in fixed row 1) ---
        tech_button_frame = ttk.Frame(tech_dialog, padding=(0, 5, 0, 5))
        tech_button_frame.grid(row=1, column=0, sticky="ew") # Place in non-expanding row 1
        # Center buttons in the frame
        tech_button_frame.columnconfigure(0, weight=1)
        tech_button_frame.columnconfigure(1, weight=1)

        def submit_action():
            tech = tech_entry.get().strip() or os.getenv('USERNAME', '')
            notes = notes_text.get("1.0", "end-1c")
            logger.info(f"Recording scan for asset {asset_id} with status {status} and site {site_value}")
            success = db.record_scan(asset_id, status, tech, notes, site_value)
            if success:
                logger.info(f"Successfully recorded scan for {asset_id}")

                # --- FIX: Show message BEFORE destroying dialogs ---
                # Use check_window as the parent, as tech_dialog is about to be destroyed indirectly
                messagebox.showinfo("Success", f"Asset {asset_id} checked {status}", parent=check_window)
                # --- END FIX ---

                tech_dialog.destroy()
                check_window.destroy()

                if callback:
                    logger.info("Executing refresh callback")
                    try:
                         callback() # Trigger full refresh
                         logger.info("Callback executed successfully")
                    except Exception as e:
                         logger.error(f"Error in callback: {str(e)}")
                # Removed messagebox call from here
            else:
                 logger.error(f"Failed to record scan for {asset_id}")
                 messagebox.showerror("Error", f"Failed to check {status} asset {asset_id}", parent=tech_dialog)

        # Use ttk buttons
        submit_btn = ttk.Button(tech_button_frame, text=f"Complete Check {status.upper()}", command=submit_action)
        submit_btn.grid(row=0, column=0, padx=5, pady=5, sticky="e")
        cancel_btn = ttk.Button(tech_button_frame, text="Cancel", command=tech_dialog.destroy)
        cancel_btn.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    # --- Button Definitions (check_in, check_out etc.) --- (Logic seems okay)
    def check_in(): record_status_change("in")
    def check_out(): record_status_change("out")
    def view_details(): check_window.destroy(); show_asset_details(db, asset_id, site_config, parent=check_window.master) # Pass parent
    def view_cmdb():
        url = asset_data.get('cmdb_url') or f"https://globalfoundries.service-now.com/u_cmdb_ci_notebook.do?sysparm_query=asset_tag%3D{asset_id}"
        webbrowser.open(url)
    def toggle_flag():
        check_window.destroy()
        if is_flagged:
            # Unflag logic (using simpledialog is okay for now, but could be a custom dialog)
             unflag_reason = simpledialog.askstring("Unflag Reason",
                                                 "Enter a reason for unflagging this asset:",
                                                 parent=check_window.master) # Pass parent
             if unflag_reason is not None: # Check if user cancelled
                 tech_name = os.getenv('USERNAME', '')
                 success = db.unflag_asset(asset_id, tech_name, unflag_reason or "Flag removed")
                 if success:
                     messagebox.showinfo("Success", f"Flag removed from asset {asset_id}", parent=check_window.master)
                     if callback: callback()
                 else:
                     messagebox.showerror("Error", f"Failed to remove flag from asset {asset_id}", parent=check_window.master)
        else:
            # Flag the asset - ensure show_flag_dialog is also scalable
            show_flag_dialog(db, asset_id, callback=callback, current_site=current_status.get('site'),
                             current_flag_status=flag_status, parent=check_window.master) # Pass parent

    # --- Bottom Button Bar (in fixed row 1) ---
    bottom_button_frame = ttk.Frame(check_window, padding=(0, 5, 0, 5))
    bottom_button_frame.grid(row=1, column=0, sticky="ew") # Place in non-expanding row 1
    # Configure columns to center the group of buttons or distribute space
    bottom_button_frame.columnconfigure(0, weight=1) # Spacer
    bottom_button_frame.columnconfigure(6, weight=1) # Spacer

    button_index = 1 # Start placing buttons from column 1

    # Conditional check-in/out buttons
    if current_state != "in":
        check_in_btn = ttk.Button(bottom_button_frame, text="Check IN", command=check_in, style="Green.TButton")
        check_in_btn.grid(row=0, column=button_index, padx=5, pady=5); button_index += 1
    if current_state != "out":
        check_out_btn = ttk.Button(bottom_button_frame, text="Check OUT", command=check_out, style="Red.TButton")
        check_out_btn.grid(row=0, column=button_index, padx=5, pady=5); button_index += 1

    # Other buttons
    details_btn = ttk.Button(bottom_button_frame, text="Details", command=view_details)
    details_btn.grid(row=0, column=button_index, padx=5, pady=5); button_index += 1
    cmdb_btn = ttk.Button(bottom_button_frame, text="CMDB", command=view_cmdb)
    cmdb_btn.grid(row=0, column=button_index, padx=5, pady=5); button_index += 1
    flag_btn = ttk.Button(bottom_button_frame, text="Unflag" if is_flagged else "Flag", command=toggle_flag, style="Orange.TButton")
    flag_btn.grid(row=0, column=button_index, padx=5, pady=5); button_index += 1
    close_btn = ttk.Button(bottom_button_frame, text="Close", command=check_window.destroy)
    close_btn.grid(row=0, column=button_index, padx=5, pady=5); button_index += 1

def show_manual_entry_form(db, identifier=None, callback=None, parent=None):
    """Show manual data entry form (Revised for Scaling)"""
    logger.info(f"Opening manual entry form for identifier: {identifier}")

    # Use create_properly_sized_dialog
    entry_window = create_properly_sized_dialog("Manual Asset Entry", min_width=500, min_height=600, parent=parent)

    # Main content frame (expandable row 0)
    main_frame = ttk.Frame(entry_window, padding="5")
    main_frame.grid(row=0, column=0, sticky="nsew")
    entry_window.rowconfigure(0, weight=1) # Ensure main frame row expands
    entry_window.columnconfigure(0, weight=1) # Ensure main frame col expands

    # Create scrollable area within main_frame
    canvas = tk.Canvas(main_frame, borderwidth=0, highlightthickness=0)
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas, padding="10") # Add padding inside

    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    add_mousewheel_scrolling(canvas, scrollable_frame) # Add mouse wheel

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    # Configure grid layout inside scrollable_frame
    scrollable_frame.columnconfigure(1, weight=1) # Allow entry column to expand

    # Field definitions (Seems okay)
    fields = [
        "asset_tag", "hostname", "serial_number", "operational_status",
        "install_status", "location", "ci_region", "owned_by",
        "assigned_to", "manufacturer", "model_id", "model_description",
        "vendor", "warranty_expiration", "os", "os_version"
    ]
    field_labels = { "asset_tag": "Asset Tag", "hostname": "Hostname", # ... rest are okay
                     "serial_number": "Serial Number", "operational_status": "Operational Status",
                     "install_status": "Install Status", "location": "Location", "ci_region": "CI Region",
                     "owned_by": "Owned By", "assigned_to": "Assigned To", "manufacturer": "Manufacturer",
                     "model_id": "Model ID", "model_description": "Model Description", "vendor": "Vendor",
                     "warranty_expiration": "Warranty Expiration", "os": "Operating System", "os_version": "OS Version" }
    entries = {}

    # Create and grid labels and entries
    for i, field in enumerate(fields):
        label = field_labels.get(field, field.replace('_', ' ').title())
        # Add asterisk for required field
        if field == "asset_tag": label += " *"
        ttk.Label(scrollable_frame, text=f"{label}:").grid(row=i, column=0, sticky="w", padx=(0, 10), pady=2)
        entry = ttk.Entry(scrollable_frame, width=50) # Keep width reasonable, rely on expansion
        entry.grid(row=i, column=1, sticky="ew", padx=0, pady=2) # sticky='ew'
        entries[field] = entry
        add_context_menu(entry)
        # Pre-fill logic (seems okay)
        is_asset = identifier and identifier.lower().startswith("gf-") if identifier else False
        if identifier and field == "asset_tag" and is_asset: entry.insert(0, identifier)
        elif identifier and field == "serial_number" and not is_asset: entry.insert(0, identifier)

    # Comments field (make it expandable)
    ttk.Label(scrollable_frame, text="Comments:").grid(row=len(fields), column=0, sticky="nw", padx=(0, 10), pady=(5,2))
    comments_text = scrolledtext.ScrolledText(scrollable_frame, width=50, height=4, wrap="word") # Use scrolledtext
    comments_text.grid(row=len(fields) + 1, column=0, columnspan=2, sticky="nsew", pady=(0,5)) # sticky='nsew'
    scrollable_frame.rowconfigure(len(fields) + 1, weight=1) # Allow comments row to expand
    add_context_menu(comments_text)

    # CMDB URL field
    ttk.Label(scrollable_frame, text="CMDB URL:").grid(row=len(fields) + 2, column=0, sticky="w", padx=(0, 10), pady=2)
    cmdb_url_entry = ttk.Entry(scrollable_frame, width=50)
    cmdb_url_entry.grid(row=len(fields) + 2, column=1, sticky="ew", padx=0, pady=(2, 5)) # sticky='ew'
    add_context_menu(cmdb_url_entry)

    # Result storage
    result = [None] # Using list to allow modification inside nested function

    # --- Save/Cancel Buttons (in fixed row 1) ---
    button_frame = ttk.Frame(entry_window, padding=(0, 5, 0, 5))
    button_frame.grid(row=1, column=0, sticky="ew")
    button_frame.columnconfigure(0, weight=1) # Center buttons
    button_frame.columnconfigure(1, weight=1)

    def save_data():
        asset_data = collect_data()
        if not asset_data: return # Error handled in collect_data
        logger.info(f"Saving manual entry data for asset: {asset_data.get('asset_tag', 'Unknown')}")
        success = db.update_asset(asset_data)
        if not success:
            logger.error(f"Failed to save asset data: {asset_data.get('asset_tag', 'Unknown')}")
            messagebox.showerror("Error", "Failed to save asset data", parent=entry_window)
            return

        logger.info(f"Successfully saved asset data: {asset_data.get('asset_tag', 'Unknown')}")
        tech_name = os.getenv('USERNAME', '')
        # Get site from config if possible (needed for initial check-in)
        # This requires passing config down or accessing it globally
        # Assuming config is available via `parent.app.config` if parent is app
        site = 'Unknown' # Default
        try:
             if parent and hasattr(parent, 'app') and hasattr(parent.app, 'config'):
                  site = parent.app.config.get('site', 'Unknown')
        except Exception:
             pass # Ignore errors getting site, use Unknown

        db.record_scan(asset_data['asset_tag'], "in", tech_name, "Initial entry via manual form", site)
        logger.info(f"Recorded initial scan for asset: {asset_data.get('asset_tag', 'Unknown')} at site {site}")

        result[0] = asset_data # Store result
        entry_window.destroy()
        if callback:
             logger.info("Executing callback after manual entry")
             try: callback()
             except Exception as e: logger.error(f"Error in callback: {str(e)}")
        return True # Indicate success

    def collect_data():
         asset_data = {field: entries[field].get().strip() for field in fields}
         asset_data['comments'] = comments_text.get("1.0", "end-1c").strip()
         asset_data['cmdb_url'] = cmdb_url_entry.get().strip()
         if not asset_data['asset_tag']:
             messagebox.showerror("Error", "Asset Tag is required", parent=entry_window)
             return None
         return asset_data

    save_btn = ttk.Button(button_frame, text="Save", command=save_data)
    save_btn.grid(row=0, column=0, padx=5, pady=5, sticky="e")
    cancel_btn = ttk.Button(button_frame, text="Cancel", command=entry_window.destroy)
    cancel_btn.grid(row=0, column=1, padx=5, pady=5, sticky="w")

    entry_window.wait_window()
    return result[0]

def edit_asset(db, asset_id, callback=None, parent=None):
    logger.info(f"Editing asset: {asset_id}")
    asset = db.get_asset_by_id(asset_id)
    if not asset: # Error handling seems okay
        logger.error(f"Asset not found: {asset_id}")
        messagebox.showerror("Error", f"Asset {asset_id} not found", parent=parent)
        return

    current_status = db.get_asset_current_status(asset_id)
    current_site = current_status.get('site') if current_status else None

    entry_window = create_properly_sized_dialog(f"Edit Asset: {asset_id}", min_width=550, min_height=650, parent=parent)

    # Main content frame (expandable row 0)
    main_frame = ttk.Frame(entry_window, padding="5")
    main_frame.grid(row=0, column=0, sticky="nsew")
    entry_window.rowconfigure(0, weight=1)
    entry_window.columnconfigure(0, weight=1)

    # Scrollable area
    canvas = tk.Canvas(main_frame, borderwidth=0, highlightthickness=0)
    scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
    scrollable_frame = ttk.Frame(canvas, padding="10")
    scrollable_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)
    add_mousewheel_scrolling(canvas, scrollable_frame)
    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")
    scrollable_frame.columnconfigure(1, weight=1) # Allow entry column to expand

    # Field definitions (Seems okay)
    fields = [ "hostname", "serial_number", "operational_status", "install_status",
               "location", "ci_region", "owned_by", "assigned_to", "manufacturer",
               "model_id", "model_description", "vendor", "warranty_expiration",
               "os", "os_version" ]
    field_labels = { "hostname": "Hostname", "serial_number": "Serial Number", # ... rest okay
                     "operational_status": "Operational Status", "install_status": "Install Status",
                     "location": "Location", "ci_region": "CI Region", "owned_by": "Owned By",
                     "assigned_to": "Assigned To", "manufacturer": "Manufacturer", "model_id": "Model ID",
                     "model_description": "Model Description", "vendor": "Vendor",
                     "warranty_expiration": "Warranty Expiration", "os": "Operating System", "os_version": "OS Version" }
    entries = {}

    # Row counter
    row_idx = 0

    # Asset Tag (Readonly)
    ttk.Label(scrollable_frame, text="Asset Tag:").grid(row=row_idx, column=0, sticky="w", padx=(0, 10), pady=2)
    asset_tag_entry = ttk.Entry(scrollable_frame, width=50)
    asset_tag_entry.grid(row=row_idx, column=1, sticky="ew", padx=0, pady=2)
    asset_tag_entry.insert(0, asset_id)
    asset_tag_entry.config(state="readonly")
    row_idx += 1

    # Editable fields
    for field in fields:
        label = field_labels.get(field, field.replace('_', ' ').title())
        ttk.Label(scrollable_frame, text=f"{label}:").grid(row=row_idx, column=0, sticky="w", padx=(0, 10), pady=2)
        entry = ttk.Entry(scrollable_frame, width=50)
        entry.grid(row=row_idx, column=1, sticky="ew", padx=0, pady=2)
        entries[field] = entry
        entry.insert(0, asset.get(field, "") or "") # Handle None
        add_context_menu(entry)
        row_idx += 1

    # Comments field (expandable)
    ttk.Label(scrollable_frame, text="Comments from CMDB:").grid(row=row_idx, column=0, sticky="nw", padx=(0, 10), pady=(5,2))
    comments_text = scrolledtext.ScrolledText(scrollable_frame, width=50, height=4, wrap="word")
    comments_text.grid(row=row_idx + 1, column=0, columnspan=2, sticky="nsew", pady=(0,5))
    scrollable_frame.rowconfigure(row_idx + 1, weight=1) # Allow comments row to expand
    comments_text.insert("1.0", asset.get('comments', "") or "")
    add_context_menu(comments_text)
    row_idx += 2

    # CMDB URL field
    ttk.Label(scrollable_frame, text="CMDB URL:").grid(row=row_idx, column=0, sticky="w", padx=(0, 10), pady=2)
    cmdb_url_entry = ttk.Entry(scrollable_frame, width=50)
    cmdb_url_entry.grid(row=row_idx, column=1, sticky="ew", padx=0, pady=(2, 5))
    cmdb_url_entry.insert(0, asset.get('cmdb_url', "") or "")
    add_context_menu(cmdb_url_entry)
    row_idx += 1

    # Current Site Display
    ttk.Label(scrollable_frame, text="Current Site:").grid(row=row_idx, column=0, sticky="w", padx=(0, 10), pady=5)
    site_display = ttk.Label(scrollable_frame, text=current_site or "Unknown")
    site_display.grid(row=row_idx, column=1, sticky="w", padx=0, pady=5)
    row_idx += 1

    # --- Action Functions (view_cmdb, sync_with_cmdb, save_data) ---
        # CMDB View function
    def view_cmdb():
        if asset.get('cmdb_url'):
            webbrowser.open(asset['cmdb_url'])
        else:
            # Generate a URL based on asset tag
            url = f"https://globalfoundries.service-now.com/u_cmdb_ci_notebook.do?sysparm_query=asset_tag%3D{asset_id}"
            webbrowser.open(url)

    def sync_with_cmdb():
         if messagebox.askyesno("Sync with CMDB",
                               "This will update this asset's data from CMDB. Any local changes not saved will be lost. Continue?",
                               parent=entry_window):
             # Indicate progress
             sync_btn.config(state="disabled", text="Syncing...")
             entry_window.update_idletasks()
             try:
                 # Use scrape_servicenow which handles user interaction
                 asset_data = scrape_servicenow(asset_id, is_asset=True) # Assuming scrape_servicenow returns data or None

                 if asset_data:
                     asset_data['asset_id'] = asset_id # Ensure asset_id is preserved
                     success = db.update_asset(asset_data)
                     if success:
                         messagebox.showinfo("Success", "Asset data updated from CMDB.", parent=entry_window)
                         entry_window.destroy() # Close current edit window
                         if callback: callback() # Refresh the main UI
                         edit_asset(db, asset_id, callback=callback, parent=parent) # Reopen edit with new data
                         return # Prevent rest of function from running
                     else:
                         messagebox.showerror("Error", "Failed to save updated asset data to database.", parent=entry_window)
                 else:
                     # scrape_servicenow likely handled showing errors/manual entry, or user cancelled
                     logger.warning("No data returned from scrape_servicenow for sync.")
                     # Optional: message if scrape didn't return data but didn't error?
                     # messagebox.showwarning("No Data", "No data retrieved from CMDB or operation cancelled.", parent=entry_window)

             except Exception as e:
                 logger.error(f"Error during CMDB sync: {e}", exc_info=True)
                 messagebox.showerror("Error", f"Failed to sync with CMDB: {str(e)}", parent=entry_window)
             finally:
                  # Re-enable button
                  sync_btn.config(state="normal", text="Sync with CMDB")

    def save_data():
        asset_data = {field: entries[field].get().strip() for field in fields}
        asset_data['comments'] = comments_text.get("1.0", "end-1c").strip()
        asset_data['cmdb_url'] = cmdb_url_entry.get().strip()
        asset_data['asset_tag'] = asset_id # Use original asset_id as PK
        logger.info(f"Saving edited data for asset: {asset_id}")
        success = db.update_asset(asset_data) # update_asset uses asset_tag/asset_id
        if not success:
            logger.error(f"Failed to update asset: {asset_id}")
            messagebox.showerror("Error", "Failed to update asset", parent=entry_window)
            return
        logger.info(f"Successfully updated asset: {asset_id}")
        db.record_scan(asset_id, "edited", os.getenv('USERNAME', ''), "Asset details edited", current_site)
        logger.info(f"Recorded edit action in history for asset: {asset_id}")
        entry_window.destroy()
        if callback:
            logger.info("Executing callback after edit")
            try: callback()
            except Exception as e: logger.error(f"Error in callback: {str(e)}")
        messagebox.showinfo("Success", f"Asset {asset_id} updated", parent=parent) # Show message on parent


    # --- Bottom Button Bar (in fixed row 1) ---
    button_frame = ttk.Frame(entry_window, padding=(0, 5, 0, 5))
    button_frame.grid(row=1, column=0, sticky="ew")
    # Distribute space around buttons
    button_frame.columnconfigure(0, weight=1)
    button_frame.columnconfigure(5, weight=1)

    save_btn = ttk.Button(button_frame, text="Save", command=save_data)
    save_btn.grid(row=0, column=1, padx=5, pady=5)
    sync_btn = ttk.Button(button_frame, text="Sync with CMDB", command=sync_with_cmdb)
    sync_btn.grid(row=0, column=2, padx=5, pady=5)
    cmdb_btn = ttk.Button(button_frame, text="View in CMDB", command=view_cmdb) # Pass asset data
    cmdb_btn.grid(row=0, column=3, padx=5, pady=5)
    cancel_btn = ttk.Button(button_frame, text="Cancel", command=entry_window.destroy)
    cancel_btn.grid(row=0, column=4, padx=5, pady=5)

def delete_asset(db, asset_id, parent_window=None, callback=None):
    logger.info(f"Attempting to delete asset: {asset_id}")
    current_status = db.get_asset_current_status(asset_id)
    current_site = current_status.get('site') if current_status else None

    # Pass parent to messagebox
    if messagebox.askyesno("Confirm Delete",
                          f"Are you sure you want to delete asset {asset_id}?\n\nThis marks the asset as deleted but keeps its history.",
                          parent=parent_window): # Pass parent

        # Soft delete logic (seems okay)
        db.record_scan(asset_id, "deleted", os.getenv('USERNAME', ''), "Asset deleted from inventory", current_site)
        success = db.delete_asset(asset_id) # delete_asset handles DB update
        if success:
            logger.info(f"Successfully deleted asset: {asset_id}")
            messagebox.showinfo("Success", f"Asset {asset_id} deleted. Scan again to restore.", parent=parent_window) # Pass parent
            if callback:
                 logger.info("Executing callback after delete")
                 try: callback()
                 except Exception as e: logger.error(f"Error in callback after delete: {str(e)}")
            if parent_window: parent_window.destroy() # Close the originating window (e.g., details window)
        else:
            logger.error(f"Failed to delete asset: {asset_id}")
            messagebox.showerror("Error", f"Failed to delete asset {asset_id}", parent=parent_window) # Pass parent

def show_asset_details(db, asset_id, site_config=None, parent=None):
    logger.info(f"Showing details for asset: {asset_id}")
    asset = db.get_asset_by_id(asset_id)
    if not asset: # Error handling okay
        logger.error(f"Asset not found: {asset_id}")
        messagebox.showerror("Error", f"Asset {asset_id} not found", parent=parent)
        return

    history = db.get_asset_history(asset_id, 5) # Recent history limit seems okay
    current_status = db.get_asset_current_status(asset_id)
    current_state = current_status.get('status', 'unknown')
    current_site = current_status.get('site')
    
    # Get both flag statuses
    flag_status = db.get_flag_status(asset_id)
    is_flagged = flag_status and flag_status.get('flag_status', False)
    
    expiry_status = db.get_expiry_flag_status(asset_id)
    is_expiring = expiry_status and expiry_status.get('expiry_flag_status', False)

    # Use create_properly_sized_dialog
    detail_window = create_properly_sized_dialog(f"Asset Details: {asset_id}", min_width=700, min_height=600, parent=parent)

    # Main container frame (expandable row 0)
    main_container = ttk.Frame(detail_window, padding="5")
    main_container.grid(row=0, column=0, sticky="nsew")
    detail_window.rowconfigure(0, weight=1)
    detail_window.columnconfigure(0, weight=1)

    # --- Tabs ---
    notebook = ttk.Notebook(main_container)
    notebook.pack(fill="both", expand=True, pady=(0, 5)) # Use pack within main_container

    # --- Tab 1: Asset Details ---
    details_tab = ttk.Frame(notebook, padding="5")
    notebook.add(details_tab, text=" Asset Details ")
    details_tab.columnconfigure(0, weight=1) # Configure grid inside tab
    details_tab.rowconfigure(1, weight=1) # Make details frame (row 1) expandable

    # Status indicator frame (non-expanding row 0)
    status_frame = ttk.Frame(details_tab)
    status_frame.grid(row=0, column=0, sticky="ew", padx=5, pady=5)
    # Status indicator contents (canvas, labels) - layout seems okay within status_frame

    # --- Status Indicator --- (Layout seems okay, added check for current_status)
    indicator_size = 24
    indicator_canvas = tk.Canvas(status_frame, width=indicator_size, height=indicator_size, highlightthickness=0)
    indicator_canvas.pack(side="left", padx=5)
    indicator_color = "#4CAF50" if current_state == "in" else "#F44336"
    indicator_canvas.create_oval(2, 2, indicator_size-2, indicator_size-2, fill=indicator_color, outline="")

    last_action_time = format_timestamp(current_status.get('timestamp', 'Unknown') if current_status else 'Unknown')
    site_display = f" at {current_site}" if current_site and current_site != 'Out' else ''
    status_label_text = f"Status: {current_state.upper()}{site_display} ({last_action_time})"
    status_label = ttk.Label(status_frame, text=status_label_text, font=("", 12, "bold"))
    status_label.pack(side="left", padx=5)
    if current_status:
        ttk.Label(status_frame, text=f"Tech: {current_status.get('tech_name', 'Unknown')}").pack(side="left", padx=20)

    # --- Flag Indicators ---
    flags_frame = ttk.Frame(status_frame)
    flags_frame.pack(side="right", padx=10)
    
    # Regular flag
    if is_flagged:
        flag_canvas = tk.Canvas(flags_frame, width=24, height=24, bg="#FF9800", highlightthickness=0)
        flag_canvas.pack(side="left", padx=(0, 5))
        flag_canvas.create_rectangle(4, 4, 8, 20, fill="#FF9800", outline="black")
        flag_canvas.create_polygon(8, 4, 20, 8, 8, 12, fill="#FF9800", outline="black")
        flag_info_text = f"Flagged by {flag_status.get('flag_tech', 'Unknown')}: {flag_status.get('flag_notes', 'No reason')[:30]}..."
        flag_info = ttk.Label(flags_frame, text=flag_info_text, foreground="#FF5722", font=("", 10, "italic"))
        flag_info.pack(side="left", padx=5)
    
    # DaaS expiry flag
    if is_expiring:
        # Use a different color and icon for expiry flag
        lease_maturity = expiry_status.get('lease_maturity_date', '')
        days_left = ""
        
        # Calculate days remaining
        if lease_maturity:
            try:
                from datetime import datetime
                
                # Try different date formats
                maturity_date = None
                for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d']:
                    try:
                        maturity_date = datetime.strptime(lease_maturity, fmt).date()
                        break
                    except (ValueError, TypeError):
                        continue
                
                if maturity_date:
                    today = datetime.now().date()
                    days_remaining = (maturity_date - today).days
                    days_left = f" ({days_remaining} days left)"
            except Exception as e:
                logger.error(f"Error calculating lease days: {e}")
        
        # Add a small separator if both flags are present
        if is_flagged:
            ttk.Separator(flags_frame, orient="vertical").pack(side="left", fill="y", padx=10, pady=5)
        
        # Create the DaaS expiry indicator
        expiry_canvas = tk.Canvas(flags_frame, width=24, height=24, bg="#9C27B0", highlightthickness=0)
        expiry_canvas.pack(side="left", padx=(0, 5))
        
        # Draw a clock-like icon
        expiry_canvas.create_oval(4, 4, 20, 20, fill="#9C27B0", outline="black")
        expiry_canvas.create_line(12, 12, 12, 6, fill="black", width=2)  # Hour hand
        expiry_canvas.create_line(12, 12, 16, 14, fill="black", width=2)  # Minute hand
        
        expiry_text = f"DaaS Expiring{days_left}"
        expiry_info = ttk.Label(flags_frame, text=expiry_text, foreground="#9C27B0", font=("", 10, "italic"))
        expiry_info.pack(side="left", padx=5)

    # --- Scrollable Details Area (expandable row 1) ---
    details_content_frame = ttk.Frame(details_tab)
    details_content_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5) # Expand this frame
    details_content_frame.rowconfigure(0, weight=1) # Canvas row expands
    details_content_frame.columnconfigure(0, weight=1) # Canvas col expands

    canvas = tk.Canvas(details_content_frame, borderwidth=0, highlightthickness=0)
    scrollbar = ttk.Scrollbar(details_content_frame, orient="vertical", command=canvas.yview)
    details_frame = ttk.Frame(canvas, padding="10") # Frame inside canvas
    canvas.create_window((0, 0), window=details_frame, anchor="nw", tags="details_frame")
    canvas.configure(yscrollcommand=scrollbar.set)
    canvas.grid(row=0, column=0, sticky="nsew") # Canvas expands
    scrollbar.grid(row=0, column=1, sticky="ns")
    details_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
    canvas.bind("<Configure>", lambda e: canvas.itemconfig("details_frame", width=e.width))
    add_mousewheel_scrolling(canvas, details_frame)
    details_frame.columnconfigure(1, weight=1) # Make value column expandable

    # --- Populate Details (inside details_frame) ---
    row = 0
    # Key filtering and label generation seems okay
    skipped_keys = ['asset_id', 'asset_tag', 'flag_status', 'flag_notes', 'flag_timestamp', 'flag_tech', 'last_updated']
    display_order = [ # Optional: Define a preferred display order
          'hostname', 'serial_number', 'manufacturer', 'model_id', 'model_description',
          'operational_status', 'install_status', 'location', 'ci_region', 'assigned_to',
          'owned_by', 'os', 'os_version', 'vendor', 'warranty_expiration', 'cmdb_url', 'comments'
     ]
    displayed_keys = set()

    # Display Asset Tag first
    ttk.Label(details_frame, text="Asset Tag:", anchor="e", width=20).grid(row=row, column=0, sticky="ne", padx=(0, 5), pady=2)
    asset_tag_entry = ttk.Entry(details_frame, width=50)
    asset_tag_entry.grid(row=row, column=1, sticky="ew", padx=5, pady=2)
    asset_tag_entry.insert(0, asset.get('asset_id', asset.get('asset_tag', ''))) # Use asset_id or asset_tag
    asset_tag_entry.config(state="readonly")
    add_context_menu(asset_tag_entry)
    row += 1

    # Display fields in preferred order
    for key in display_order:
         if key in asset and key not in skipped_keys:
              field_name = " ".join(word.capitalize() for word in key.split('_'))
              ttk.Label(details_frame, text=f"{field_name}:", anchor="e", width=20).grid(row=row, column=0, sticky="ne", padx=(0, 5), pady=2)
              value = asset.get(key, "") or "" # Handle None

              # Use Text widget for potentially long fields (like comments, URL)
              if key in ['comments', 'cmdb_url'] or (isinstance(value, str) and len(value) > 60):
                   field_text = Text(details_frame, height=1, width=50, wrap="none", font=("", 9), borderwidth=0, relief="flat")
                   # Adjust height for comments
                   if key == 'comments': field_text.config(height=3, wrap="word")
                   field_text.grid(row=row, column=1, sticky="ew", padx=5, pady=2)
                   field_text.insert("1.0", str(value))
                   field_text.config(state="disabled")
                   add_context_menu(field_text) # Add context menu to Text widget
              else:
                   # Use readonly Entry for shorter fields
                   entry = ttk.Entry(details_frame, width=50)
                   entry.grid(row=row, column=1, sticky="ew", padx=5, pady=2)
                   entry.insert(0, str(value))
                   entry.config(state="readonly")
                   add_context_menu(entry)
              displayed_keys.add(key)
              row += 1

     # Display any remaining fields not in the preferred order
    for key, value in asset.items():
         if key not in skipped_keys and key not in displayed_keys:
              field_name = " ".join(word.capitalize() for word in key.split('_'))
              ttk.Label(details_frame, text=f"{field_name}:", anchor="e", width=20).grid(row=row, column=0, sticky="ne", padx=(0, 5), pady=2)
              value = value or ""
              entry = ttk.Entry(details_frame, width=50)
              entry.grid(row=row, column=1, sticky="ew", padx=5, pady=2)
              entry.insert(0, str(value))
              entry.config(state="readonly")
              add_context_menu(entry)
              row += 1


    # --- Notes Section --- (Appears okay, but ensure parent frame resizes)
    notes_label = ttk.Label(details_frame, text="Recent Notes:", anchor="e", width=20)
    notes_label.grid(row=row, column=0, sticky="ne", padx=(0, 5), pady=(10, 2))
    notes_frame = ttk.Frame(details_frame)
    notes_frame.grid(row=row + 1, column=0, columnspan=2, sticky="nsew", pady=(0,5))
    details_frame.rowconfigure(row + 1, weight=1) # Allow notes frame row to expand
    notes_frame.rowconfigure(0, weight=1)
    notes_frame.columnconfigure(0, weight=1)
    notes_text = scrolledtext.ScrolledText(notes_frame, height=5, wrap="word", font=("", 9)) # Use scrolledtext
    notes_text.grid(row=0, column=0, sticky="nsew")
    add_context_menu(notes_text)
    # Populate notes (logic seems okay)
    notes_text.config(state="normal")
    notes_text.delete("1.0", "end")
    notes_history = [item for item in history if item.get('notes') and item['notes'].strip() and not item['notes'].startswith("Initial entry")]
    if notes_history:
         for item in notes_history:
              ts = format_timestamp(item.get('timestamp', '')) # Use consistent formatting
              tech = item.get('tech_name', '')
              status = item.get('status', '').upper()
              note = item.get('notes', '')
              notes_text.insert("end", f"[{ts}] {tech} ({status}): {note}\n\n")
    else:
         notes_text.insert("end", "No recent notes available.")
    notes_text.config(state="disabled")


    # --- Tab 2: History ---
    history_tab = ttk.Frame(notebook, padding="10")
    notebook.add(history_tab, text=" Recent History ")
    history_tab.rowconfigure(0, weight=1) # Make treeview row expandable
    history_tab.columnconfigure(0, weight=1) # Make treeview col expandable

    history_frame = ttk.Frame(history_tab)
    history_frame.grid(row=0, column=0, sticky="nsew") # History frame expands
    history_frame.rowconfigure(0, weight=1)
    history_frame.columnconfigure(0, weight=1)

    # History treeview columns/setup seems okay
    columns = ("timestamp", "status", "tech", "notes", "site") # Added site
    history_tree = ttk.Treeview(history_frame, columns=columns, show="headings")
    history_tree.heading("timestamp", text="Date/Time")
    history_tree.heading("status", text="Action")
    history_tree.heading("tech", text="Technician")
    history_tree.heading("notes", text="Notes")
    history_tree.heading("site", text="Site") # Added site heading
    history_tree.column("timestamp", width=140, stretch=False)
    history_tree.column("status", width=80, stretch=False)
    history_tree.column("tech", width=100, stretch=False)
    history_tree.column("notes", width=250) # Allow notes to stretch
    history_tree.column("site", width=80, stretch=False) # Added site column

    history_scrollbar = ttk.Scrollbar(history_frame, orient="vertical", command=history_tree.yview)
    history_tree.configure(yscrollcommand=history_scrollbar.set)
    history_tree.grid(row=0, column=0, sticky="nsew") # Treeview expands
    history_scrollbar.grid(row=0, column=1, sticky="ns")
    add_context_menu(history_tree) # Add context menu

    # Populate history (logic seems okay, added site)
    for item in history:
         tags = ()
         if item.get('status') in ['flagged', 'unflagged']: tags = ('flag_action',)
         history_tree.insert("", "end", values=(
             format_timestamp(item.get('timestamp', '')),
             item.get('status', ''),
             item.get('tech_name', ''),
             item.get('notes', ''),
             item.get('site', '') # Added site value
         ), tags=tags)
    history_tree.tag_configure('flag_action', background='#FFF9C4') # Lighter yellow/orange


    # --- Action Functions (refresh_and_close, add_new_note, toggle_flag, open_cmdb) ---
    # Ensure dialogs opened by these are also scalable
    # Add Note Dialog:
    def add_new_note():
         note_dialog = create_properly_sized_dialog(f"Add Note to {asset_id}", min_width=400, min_height=300, parent=detail_window)
         note_content = ttk.Frame(note_dialog, padding="10")
         note_content.grid(row=0, column=0, sticky="nsew")
         note_dialog.rowconfigure(0, weight=1); note_dialog.columnconfigure(0, weight=1) # Configure dialog grid
         note_content.rowconfigure(1, weight=1) # Text widget row expands
         note_content.columnconfigure(0, weight=1) # Text widget col expands

         ttk.Label(note_content, text="New Note:").grid(row=0, column=0, sticky="w", pady=(0,2))
         new_note_text = scrolledtext.ScrolledText(note_content, width=50, height=8, wrap="word")
         new_note_text.grid(row=1, column=0, sticky="nsew", pady=(0, 5))
         add_context_menu(new_note_text)

         tech_frame = ttk.Frame(note_content)
         tech_frame.grid(row=2, column=0, sticky="ew", pady=(5,0))
         ttk.Label(tech_frame, text="Technician:").pack(side="left")
         tech_entry = ttk.Entry(tech_frame, width=25)
         tech_entry.pack(side="left", padx=5, fill='x', expand=True)
         tech_entry.insert(0, os.getenv('USERNAME', ''))
         add_context_menu(tech_entry)
         new_note_text.focus()

         note_button_frame = ttk.Frame(note_dialog, padding=(0, 5, 0, 5))
         note_button_frame.grid(row=1, column=0, sticky="ew")
         note_button_frame.columnconfigure(0, weight=1); note_button_frame.columnconfigure(1, weight=1) # Center buttons

         def save_note():
              note = new_note_text.get("1.0", "end-1c").strip()
              tech = tech_entry.get().strip() or os.getenv('USERNAME', '')
              if not note:
                  messagebox.showwarning("Empty Note", "Please enter a note.", parent=note_dialog)
                  return
              success = db.record_scan(asset_id, "note", tech, note, current_site) # Use current_site
              if success:
                  note_dialog.destroy()
                  refresh_and_close() # Refresh main details view
              else: messagebox.showerror("Error", "Failed to save note.", parent=note_dialog)

         save_btn = ttk.Button(note_button_frame, text="Save Note", command=save_note)
         save_btn.grid(row=0, column=0, sticky='e', padx=5)
         cancel_btn = ttk.Button(note_button_frame, text="Cancel", command=note_dialog.destroy)
         cancel_btn.grid(row=0, column=1, sticky='w', padx=5)

    # Refresh function
    def refresh_and_close():
         detail_window.destroy()
         # Reopen the details window to show updated info
         show_asset_details(db, asset_id, site_config, parent=parent)

    # Toggle Flag function (ensure show_flag_dialog is scalable)
    def toggle_flag():
        # Existing logic using simpledialog or calling show_flag_dialog seems okay,
        # but ensure show_flag_dialog itself is scalable.
        if is_flagged:
             if messagebox.askyesno("Confirm Unflag", f"Remove the flag from asset {asset_id}?", parent=detail_window):
                 unflag_reason = simpledialog.askstring("Unflag Reason", "Reason for unflagging:", parent=detail_window)
                 if unflag_reason is not None:
                     tech_name = os.getenv('USERNAME', '')
                     success = db.unflag_asset(asset_id, tech_name, unflag_reason or "Flag removed")
                     if success:
                          messagebox.showinfo("Success", "Flag removed.", parent=detail_window)
                          refresh_and_close()
                     else: messagebox.showerror("Error", "Failed to remove flag.", parent=detail_window)
        else:
             show_flag_dialog(db, asset_id, callback=refresh_and_close,
                              current_site=current_site, current_flag_status=flag_status, parent=detail_window)

    # --- Bottom Action Buttons (fixed row 1) ---
    action_frame = ttk.Frame(detail_window, padding=(0, 5, 0, 5))
    action_frame.grid(row=1, column=0, sticky="ew")
    # Configure columns to distribute space or center buttons
    action_frame.columnconfigure(0, weight=1) # Spacer left
    action_frame.columnconfigure(8, weight=1) # Spacer right

    def open_cmdb(asset_data):
        if asset_data.get('cmdb_url'):
            webbrowser.open(asset_data['cmdb_url'])
        else:
            # Generate a URL based on asset tag
            url = f"https://globalfoundries.service-now.com/u_cmdb_ci_notebook.do?sysparm_query=asset_tag%3D{asset_id}"
            webbrowser.open(url)

    def display_asset_details(self, asset_data, current_status=None):
        """Display asset details in the main window, now including lease information"""
        self.details_text.config(state="normal")
        self.details_text.delete("1.0", "end")
        
        # Basic asset details
        details = f"""Asset Tag: {asset_data.get('asset_tag', '')}
    Serial Number: {asset_data.get('serial_number', '')}
    Hostname: {asset_data.get('hostname', '')}
    Make/Model: {asset_data.get('manufacturer', '')} {asset_data.get('model_description', '')}
    Assigned To: {asset_data.get('assigned_to', '')}
    Location: {asset_data.get('location', '')}
    OS: {asset_data.get('os', '')} {asset_data.get('os_version', '')}
    Warranty Expires: {asset_data.get('warranty_expiration', '')}
    """
        
        # Add lease information if available
        lease_start = asset_data.get('lease_start_date', '')
        lease_end = asset_data.get('lease_maturity_date', '')
        
        if lease_start or lease_end:
            details += f"\nLease Information:\n"
            if lease_start:
                details += f"Lease Start Date: {lease_start}\n"
            if lease_end:
                details += f"Lease Maturity Date: {lease_end}\n"
                
                # Calculate days remaining if lease end date exists
                try:
                    from datetime import datetime
                    
                    # Try different date formats
                    maturity_date = None
                    for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d']:
                        try:
                            maturity_date = datetime.strptime(lease_end, fmt).date()
                            break
                        except (ValueError, TypeError):
                            continue
                    
                    if maturity_date:
                        today = datetime.now().date()
                        days_remaining = (maturity_date - today).days
                        
                        if days_remaining < 0:
                            details += f"Lease Status: EXPIRED ({abs(days_remaining)} days ago)\n"
                        elif days_remaining <= 90:
                            details += f"Lease Status: EXPIRING SOON ({days_remaining} days remaining)\n"
                        else:
                            details += f"Lease Status: Active ({days_remaining} days remaining)\n"
                except Exception as e:
                    logger.error(f"Error calculating lease days: {e}")
        
        # Add current status information if available
        if current_status:
            status_text = current_status.get('status', 'unknown')
            timestamp = current_status.get('timestamp', '')
            tech_name = current_status.get('tech_name', '')
            site = current_status.get('site', '')
            
            details += f"""
    Current Status: {status_text.upper()}
    Last Action: {timestamp}
    Technician: {tech_name}
    """
            if site:
                details += f"Site: {site}\n"
        
        # Add site information from config
        if self.config and self.config.get('site'):
            details += f"\nCurrent Site: {self.config.get('site')}\n"
        
        # Add comments
        details += f"\nComments: {asset_data.get('comments', '')}"
        
        self.details_text.insert("1.0", details)
        self.details_text.config(state="disabled")

    # Use styles defined earlier
    check_in_btn = ttk.Button(action_frame, text="Check In", style="Green.TButton", state="disabled" if current_state == "in" else "normal",
                            command=lambda: show_check_in_out_dialog(db, asset, default_status="in",
                                                                   callback=refresh_and_close, site_config=site_config))
    check_in_btn.grid(row=0, column=1, padx=3, pady=5)

    check_out_btn = ttk.Button(action_frame, text="Check Out", style="Red.TButton", state="disabled" if current_state == "out" else "normal",
                             command=lambda: show_check_in_out_dialog(db, asset, default_status="out",
                                                                    callback=refresh_and_close, site_config=site_config))
    check_out_btn.grid(row=0, column=2, padx=3, pady=5)

    flag_btn = ttk.Button(action_frame, text="Unflag" if is_flagged else "Flag", style="Orange.TButton", command=toggle_flag)
    flag_btn.grid(row=0, column=3, padx=3, pady=5)

    note_btn = ttk.Button(action_frame, text="Add Note", command=add_new_note)
    note_btn.grid(row=0, column=4, padx=3, pady=5)
    edit_btn = ttk.Button(action_frame, text="Edit", command=lambda: edit_asset(db, asset_id, callback=refresh_and_close, parent=detail_window))
    edit_btn.grid(row=0, column=5, padx=3, pady=5)
    cmdb_btn = ttk.Button(action_frame, text="View CMDB", command=lambda: open_cmdb(asset))
    cmdb_btn.grid(row=0, column=6, padx=3, pady=5)
    # Delete button should perhaps be less prominent or styled differently
    delete_btn = ttk.Button(action_frame, text="Delete", command=lambda: delete_asset(db, asset_id, detail_window, callback=refresh_and_close))
    delete_btn.grid(row=0, column=7, padx=3, pady=5)
    close_btn = ttk.Button(action_frame, text="Close", command=detail_window.destroy)
    close_btn.grid(row=0, column=8, padx=(15, 5), pady=5, sticky='e') # Push close to right

# --- show_flag_dialog ---
def show_flag_dialog(db, asset_id, callback=None, current_site=None, current_flag_status=None, parent=None):
    logger.info(f"Opening flag dialog for asset: {asset_id}")

    flag_window = create_properly_sized_dialog("Flag Asset", min_width=450, min_height=300, parent=parent)

    # Content frame (expandable row 0)
    content_frame = ttk.Frame(flag_window, padding="10")
    content_frame.grid(row=0, column=0, sticky="nsew")
    flag_window.rowconfigure(0, weight=1); flag_window.columnconfigure(0, weight=1) # Dialog grid
    content_frame.rowconfigure(2, weight=1) # Notes row expands
    content_frame.columnconfigure(0, weight=1) # Notes col expands


    ttk.Label(content_frame, text=f"Asset ID: {asset_id}", font=("", 12, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 10))
    ttk.Label(content_frame, text="Reason for flagging (Required):").grid(row=1, column=0, sticky="nw", pady=(0, 2))
    notes_text = scrolledtext.ScrolledText(content_frame, width=50, height=6, wrap="word")
    notes_text.grid(row=2, column=0, sticky="nsew", pady=(0, 5))
    add_context_menu(notes_text)
    if current_flag_status and current_flag_status.get('flag_notes'): # Pre-fill logic okay
        notes_text.insert("1.0", current_flag_status.get('flag_notes', ''))
    notes_text.focus()

    tech_frame = ttk.Frame(content_frame)
    tech_frame.grid(row=3, column=0, sticky="ew", pady=(5,0))
    ttk.Label(tech_frame, text="Technician:").pack(side="left")
    tech_entry = ttk.Entry(tech_frame, width=30)
    tech_entry.pack(side="left", padx=5, fill='x', expand=True)
    tech_entry.insert(0, os.getenv('USERNAME', ''))
    add_context_menu(tech_entry)

    # Buttons (fixed row 1)
    button_frame = ttk.Frame(flag_window, padding=(0, 5, 0, 5))
    button_frame.grid(row=1, column=0, sticky="ew")
    button_frame.columnconfigure(0, weight=1); button_frame.columnconfigure(1, weight=1) # Center buttons

    def save_flag():
        flag_notes = notes_text.get("1.0", "end-1c").strip()
        tech = tech_entry.get().strip() or os.getenv('USERNAME', '')
        if not flag_notes:
            messagebox.showwarning("Missing Information", "Please enter a reason for flagging.", parent=flag_window)
            return
        success = db.flag_asset(asset_id, flag_notes, tech) # Flagging logic okay
        if success:
            messagebox.showinfo("Success", f"Asset {asset_id} flagged.", parent=flag_window)
            flag_window.destroy()
            if callback: callback()
        else: messagebox.showerror("Error", f"Failed to flag asset {asset_id}", parent=flag_window)

    # Use ttk buttons and style
    save_btn = ttk.Button(button_frame, text="Flag Asset", command=save_flag, style="Orange.TButton")
    save_btn.grid(row=0, column=0, padx=5, sticky='e')
    cancel_btn = ttk.Button(button_frame, text="Cancel", command=flag_window.destroy)
    cancel_btn.grid(row=0, column=1, padx=5, sticky='w')

class AuditDialog(Toplevel):
    """
    Enhanced dialog for performing a bench audit using Tkinter.
    Includes multiple result categories and interaction features.
    """
    def __init__(self, parent, db_manager, site_id, config): # Added config
        super().__init__(parent)
        self.db_manager = db_manager
        self.site_id = site_id
        self.config = config # Store config
        self.title(f"Audit Bench for Site: {self.site_id}")

        # Set minimum size and allow resizing
        min_width = 950
        min_height = 650
        self.minsize(min_width, min_height)
        self.resizable(True, True)

        # Center window based on minimum size
        self.update_idletasks()
        screen_width = self.winfo_screenwidth()
        screen_height = self.winfo_screenheight()
        x = (screen_width - min_width) // 2
        y = (screen_height - min_height) // 2
        self.geometry(f"{min_width}x{min_height}+{x}+{y}") # Start at min size

        # Configure main grid layout
        self.rowconfigure(0, weight=1) # Main frame expands
        self.rowconfigure(1, weight=0) # Button frame fixed
        self.columnconfigure(0, weight=1)

        # --- Main Frame ---
        # This frame goes into the expandable row 0
        main_frame = ttk.Frame(self, padding="10")
        main_frame.grid(row=0, column=0, sticky="nsew")
        # Configure main_frame grid (Input area fixed height, Results area expands)
        main_frame.rowconfigure(0, weight=0) # Input Frame
        main_frame.rowconfigure(1, weight=1) # Results Frame
        main_frame.columnconfigure(0, weight=1)


        # --- Input Area ---
        input_frame = ttk.LabelFrame(main_frame, text="Scan or Enter Asset Tags/Serial Numbers (one per line)", padding="5")
        # Place input_frame in row 0, make it fill horizontally but not expand vertically
        input_frame.grid(row=0, column=0, sticky="ew", pady=(0, 10))
        input_frame.columnconfigure(0, weight=1) # Allow text area inside to expand
        input_frame.rowconfigure(0, weight=0) # Text area fixed height

        self.asset_input_text = Text(input_frame, height=6, width=50) # Fixed height for input
        input_scrollbar = ttk.Scrollbar(input_frame, orient=tk.VERTICAL, command=self.asset_input_text.yview)
        self.asset_input_text['yscrollcommand'] = input_scrollbar.set
        # Use grid within input_frame for Text and Scrollbar
        self.asset_input_text.grid(row=0, column=0, sticky="ew")
        input_scrollbar.grid(row=0, column=1, sticky="ns")
        add_context_menu(self.asset_input_text) # Add context menu


        # --- Results Area (Now 6 columns) ---
        # Place results_frame in row 1, make it fill available space
        results_frame = ttk.Frame(main_frame)
        results_frame.grid(row=1, column=0, sticky="nsew", pady=(5,0))
        # Configure columns to have equal weight for horizontal expansion
        results_frame.columnconfigure(0, weight=1)
        results_frame.columnconfigure(1, weight=1)
        results_frame.columnconfigure(2, weight=1)
        # Configure rows: Labels fixed height, Listboxes expand vertically
        results_frame.rowconfigure(0, weight=0) # Row 1 Labels
        results_frame.rowconfigure(1, weight=1) # Row 1 Listboxes
        results_frame.rowconfigure(2, weight=0) # Row 2 Labels
        results_frame.rowconfigure(3, weight=1) # Row 2 Listboxes


        # --- Row 1 of Results ---
        ttk.Label(results_frame, text="Matching").grid(row=0, column=0, padx=5, pady=(0,2), sticky=tk.W)
        ttk.Label(results_frame, text="Missing from Bench").grid(row=0, column=1, padx=5, pady=(0,2), sticky=tk.W)
        ttk.Label(results_frame, text="Wrong Bench").grid(row=0, column=2, padx=5, pady=(0,2), sticky=tk.W)

        self.matching_list_ui = self._create_listbox_with_scrollbar(results_frame)
        self.missing_list_ui = self._create_listbox_with_scrollbar(results_frame)
        self.wrong_bench_list_ui = self._create_listbox_with_scrollbar(results_frame)

        # Place listboxes in expanding row 1
        self.matching_list_ui['frame'].grid(row=1, column=0, padx=5, pady=(0, 5), sticky=tk.NSEW)
        self.missing_list_ui['frame'].grid(row=1, column=1, padx=5, pady=(0, 5), sticky=tk.NSEW)
        self.wrong_bench_list_ui['frame'].grid(row=1, column=2, padx=5, pady=(0, 5), sticky=tk.NSEW)

        # --- Row 2 of Results ---
        ttk.Label(results_frame, text="Checked Out").grid(row=2, column=0, padx=5, pady=(0,2), sticky=tk.W)
        ttk.Label(results_frame, text="FLAGGED!").grid(row=2, column=1, padx=5, pady=(0,2), sticky=tk.W)
        ttk.Label(results_frame, text="Not in Database").grid(row=2, column=2, padx=5, pady=(0,2), sticky=tk.W)

        self.checked_out_list_ui = self._create_listbox_with_scrollbar(results_frame)
        self.flagged_list_ui = self._create_listbox_with_scrollbar(results_frame)
        self.not_in_db_list_ui = self._create_listbox_with_scrollbar(results_frame)

        # Place listboxes in expanding row 3
        self.checked_out_list_ui['frame'].grid(row=3, column=0, padx=5, sticky=tk.NSEW)
        self.flagged_list_ui['frame'].grid(row=3, column=1, padx=5, sticky=tk.NSEW)
        self.not_in_db_list_ui['frame'].grid(row=3, column=2, padx=5, sticky=tk.NSEW)

        # --- Bind Events ---
        self._bind_listbox_events(self.matching_list_ui['listbox'], 'matching')
        self._bind_listbox_events(self.missing_list_ui['listbox'], 'missing')
        self._bind_listbox_events(self.wrong_bench_list_ui['listbox'], 'wrong_bench')
        self._bind_listbox_events(self.checked_out_list_ui['listbox'], 'checked_out')
        self._bind_listbox_events(self.flagged_list_ui['listbox'], 'flagged')
        self._bind_listbox_events(self.not_in_db_list_ui['listbox'], 'not_in_db')


        # --- Buttons ---
        # Place button frame in the fixed row 1 at the bottom of the main dialog grid
        button_frame = ttk.Frame(self) # Attach directly to self (the Toplevel window)
        button_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=(5, 10)) # Use grid placement

        self.status_var = tk.StringVar(value="Enter assets and click 'Start Audit'")
        # Use grid within button_frame
        status_label = ttk.Label(button_frame, textvariable=self.status_var, relief=tk.SUNKEN, anchor=tk.W)
        status_label.grid(row=0, column=0, sticky="ew", padx=5)

        audit_button = ttk.Button(button_frame, text="Start Audit", command=self.perform_audit)
        audit_button.grid(row=0, column=1, padx=5)

        close_button = ttk.Button(button_frame, text="Close", command=self.destroy)
        close_button.grid(row=0, column=2, padx=5)

        # Configure button_frame grid weights
        button_frame.columnconfigure(0, weight=1) # Status label expands
        button_frame.columnconfigure(1, weight=0) # Buttons fixed size
        button_frame.columnconfigure(2, weight=0)

        # Focus on input area initially
        self.asset_input_text.focus_set()
        self.wait_visibility()
        self.grab_set()


    def _create_listbox_with_scrollbar(self, parent):
        """Helper to create a Listbox with a vertical Scrollbar that expands."""
        frame = ttk.Frame(parent)
        # Configure frame's grid for listbox and scrollbar
        frame.rowconfigure(0, weight=1)
        frame.columnconfigure(0, weight=1)

        scrollbar = ttk.Scrollbar(frame, orient=tk.VERTICAL)
        listbox = Listbox(frame, yscrollcommand=scrollbar.set, exportselection=False)
        scrollbar.config(command=listbox.yview)

        # Place using grid within the helper frame
        listbox.grid(row=0, column=0, sticky="nsew")
        scrollbar.grid(row=0, column=1, sticky="ns")

        add_context_menu(listbox) # Add context menu
        return {'frame': frame, 'listbox': listbox}

    # --- (_clear_all_listboxes, perform_audit methods - logic unchanged) ---
    def _clear_all_listboxes(self):
        """Clears all result listboxes."""
        self.matching_list_ui['listbox'].delete(0, tk.END)
        self.missing_list_ui['listbox'].delete(0, tk.END)
        self.wrong_bench_list_ui['listbox'].delete(0, tk.END)
        self.checked_out_list_ui['listbox'].delete(0, tk.END)
        self.flagged_list_ui['listbox'].delete(0, tk.END)
        self.not_in_db_list_ui['listbox'].delete(0, tk.END)

    def perform_audit(self):
        """Fetches data, performs comparison based on defined logic, and displays results."""
        self._clear_all_listboxes()
        self.status_var.set("Performing audit...")
        self.update_idletasks() # Update UI to show status change

        # --- (Data fetching and categorization logic - unchanged) ---
        scanned_assets_raw = self.asset_input_text.get("1.0", tk.END).splitlines()
        scanned_assets = {asset.strip().upper() for asset in scanned_assets_raw if asset.strip()}
        if not scanned_assets:
            messagebox.showwarning("Input Required", "Please enter or scan at least one asset tag/serial number.", parent=self)
            self.status_var.set("Ready.")
            return
        try:
            all_checked_in = self.db_manager.get_current_inventory(include_deleted=False)
            expected_bench_assets = {
                asset['asset_id'].upper(): asset for asset in all_checked_in
                if asset.get('site') == self.site_id and asset.get('asset_id')
            }
            expected_bench_ids = set(expected_bench_assets.keys())

            matching, wrong_bench, checked_out, flagged, not_in_db = [], [], [], [], []
            found_in_db_ids = set()

            for asset_id_upper in scanned_assets:
                asset_info = self.db_manager.get_asset_by_id(asset_id_upper)
                if not asset_info: asset_info = self.db_manager.get_asset_by_serial(asset_id_upper)
                if not asset_info:
                    not_in_db.append(asset_id_upper)
                    continue

                actual_asset_id = asset_info['asset_id']
                found_in_db_ids.add(actual_asset_id.upper())
                flag_info = self.db_manager.get_flag_status(actual_asset_id)
                if flag_info and flag_info.get('flag_status'):
                    flagged.append(actual_asset_id)
                    continue

                status_info = self.db_manager.get_asset_current_status(actual_asset_id)
                current_status = status_info.get('status', 'unknown')
                current_site = status_info.get('site')

                if current_status == 'out': checked_out.append(actual_asset_id)
                elif current_status == 'in':
                    if current_site == self.site_id: matching.append(actual_asset_id)
                    else: wrong_bench.append(f"{actual_asset_id} (Site: {current_site or 'Unknown'})")
                else: wrong_bench.append(f"{actual_asset_id} (Status: {current_status})")

            missing = sorted(list(expected_bench_ids - found_in_db_ids))

            # --- Populate Lists ---
            for asset in sorted(matching): self.matching_list_ui['listbox'].insert(tk.END, asset)
            for asset in sorted(missing): self.missing_list_ui['listbox'].insert(tk.END, asset)
            for asset in sorted(wrong_bench): self.wrong_bench_list_ui['listbox'].insert(tk.END, asset)
            for asset in sorted(checked_out): self.checked_out_list_ui['listbox'].insert(tk.END, asset)
            for asset in sorted(flagged): self.flagged_list_ui['listbox'].insert(tk.END, asset); self.flagged_list_ui['listbox'].itemconfig(tk.END, {'bg':'#FFF3E0'}) # Highlight flagged
            for asset in sorted(not_in_db): self.not_in_db_list_ui['listbox'].insert(tk.END, asset)

            msg = (f"Audit Complete. Matching: {len(matching)}, Missing: {len(missing)}, Wrong Bench: {len(wrong_bench)}, "
                   f"Checked Out: {len(checked_out)}, FLAGGED: {len(flagged)}, Not in DB: {len(not_in_db)}")
            self.status_var.set("Audit complete.")
            messagebox.showinfo("Audit Complete", msg.replace(", ", "\n"), parent=self) # Use newline for better readability

        except Exception as e:
            logger.error(f"Error during audit process: {e}", exc_info=True)
            messagebox.showerror("Audit Error", f"An error occurred during the audit:\n{e}", parent=self)
            self.status_var.set("Error during audit.")

    # --- (_bind_listbox_events, _get_selected_asset_id, _on_listbox_double_click methods - logic unchanged) ---
    def _bind_listbox_events(self, listbox_widget, category_name):
        """Binds double-click and right-click events to a listbox."""
        listbox_widget.bind("<Double-Button-1>", lambda event, cat=category_name: self._on_listbox_double_click(event, cat))
        listbox_widget.bind("<Button-3>", lambda event, cat=category_name: self._show_context_menu(event, cat))

    def _get_selected_asset_id(self, listbox_widget):
        """Gets the asset ID from the selected item in a listbox."""
        selection_indices = listbox_widget.curselection()
        if not selection_indices: return None
        selected_text = listbox_widget.get(selection_indices[0])
        asset_id = selected_text.split(" ")[0] # Handle cases like "ASSETID (Site: X)"
        return asset_id

    def _on_listbox_double_click(self, event, category_name):
        """Handles double-click event on a listbox item."""
        listbox_widget = event.widget
        asset_id = self._get_selected_asset_id(listbox_widget)
        if not asset_id: return

        logger.info(f"Double-click on '{asset_id}' in category '{category_name}'")
        if category_name == 'not_in_db':
            input_type = simpledialog.askstring("Lookup Type", f"Is '{asset_id}' an Asset Tag or Serial Number?", initialvalue="asset", parent=self)
            if input_type and input_type.lower() in ['asset', 'serial']:
                 trigger_servicenow_lookup(self.db_manager, asset_id, self, self.config, is_asset=(input_type.lower() == 'asset'))
            elif input_type: messagebox.showwarning("Invalid Type", "Please enter 'asset' or 'serial'.", parent=self)
        else:
            try: show_asset_details(self.db_manager, asset_id, self.config, parent=self)
            except Exception as e: logger.error(f"Error showing details for {asset_id}: {e}", exc_info=True); messagebox.showerror("Error", f"Could not display details for {asset_id}:\n{e}", parent=self)

    # --- (_show_context_menu, _copy_to_clipboard, _trigger_servicenow_from_menu methods - logic unchanged) ---
    def _show_context_menu(self, event, category_name):
        """Shows a right-click context menu for a listbox item."""
        listbox_widget = event.widget
        item_index = listbox_widget.nearest(event.y)
        listbox_widget.selection_clear(0, tk.END)
        listbox_widget.selection_set(item_index)
        listbox_widget.activate(item_index)

        asset_id = self._get_selected_asset_id(listbox_widget)
        if not asset_id: return

        menu = Menu(listbox_widget, tearoff=0)
        if category_name == 'not_in_db':
            menu.add_command(label=f"Lookup '{asset_id}' in ServiceNow", command=lambda id=asset_id: self._trigger_servicenow_from_menu(id))
        else:
            menu.add_command(label=f"View Details for '{asset_id}'", command=lambda id=asset_id: show_asset_details(self.db_manager, id, self.config, parent=self))
        menu.add_separator()
        menu.add_command(label=f"Copy '{asset_id}'", command=lambda id=asset_id: self._copy_to_clipboard(id))
        try: menu.tk_popup(event.x_root, event.y_root)
        finally: menu.grab_release()

    def _copy_to_clipboard(self, text_to_copy):
        """Copies the given text to the clipboard."""
        self.clipboard_clear()
        self.clipboard_append(text_to_copy)
        logger.info(f"Copied '{text_to_copy}' to clipboard.")
        self.status_var.set(f"Copied '{text_to_copy}' to clipboard.")

    def _trigger_servicenow_from_menu(self, asset_id):
         """Helper to call ServiceNow lookup from context menu."""
         input_type = simpledialog.askstring("Lookup Type", f"Is '{asset_id}' an Asset Tag or Serial Number?", initialvalue="asset", parent=self)
         if input_type and input_type.lower() in ['asset', 'serial']:
              trigger_servicenow_lookup(self.db_manager, asset_id, self, self.config, is_asset=(input_type.lower() == 'asset'))
         elif input_type: messagebox.showwarning("Invalid Type", "Please enter 'asset' or 'serial'.", parent=self)

# --- ServiceNow Lookup Trigger (logic unchanged) ---
def trigger_servicenow_lookup(db, identifier, parent_dialog, config, is_asset):
    """Initiates ServiceNow scraping in a separate thread."""
    logger.info(f"Starting ServiceNow lookup for: {identifier} (is_asset={is_asset})")
    parent_dialog.status_var.set(f"Fetching {identifier} from ServiceNow...")
    def scrape_thread():
        result_data, error_msg = None, None
        try: result_data = scrape_servicenow(identifier, is_asset)
        except Exception as ex: error_msg = str(ex); logger.error(f"Error scraping ServiceNow for {identifier}: {error_msg}", exc_info=True)
        parent_dialog.after(100, lambda data=result_data, err=error_msg: _handle_scrape_result_for_audit(data, err, db, parent_dialog, config, identifier))
    threading.Thread(target=scrape_thread, daemon=True).start()

# --- ServiceNow Result Handler (logic unchanged) ---
def _handle_scrape_result_for_audit(data, error_msg, db, parent_dialog, config, original_identifier):
    """Handles the result of the ServiceNow scrape within the audit context."""
    if error_msg: messagebox.showerror("ServiceNow Error", f"Failed to retrieve data for {original_identifier}:\n{error_msg}", parent=parent_dialog); parent_dialog.status_var.set(f"ServiceNow lookup failed."); return
    if not data: messagebox.showwarning("Not Found", f"Asset {original_identifier} not found in ServiceNow.", parent=parent_dialog); parent_dialog.status_var.set(f"Asset {original_identifier} not found."); return

    if 'asset_tag' not in data and 'asset_id' in data: data['asset_tag'] = data['asset_id']
    elif 'asset_tag' not in data: logger.error("Scraped data missing 'asset_tag'/'asset_id'"); messagebox.showerror("Data Error", "Scraped data missing Asset Tag.", parent=parent_dialog); parent_dialog.status_var.set("Error processing data."); return

    asset_id_scraped = data['asset_tag']
    logger.info(f"Successfully scraped data for {original_identifier}, found asset tag: {asset_id_scraped}")
    success = db.update_asset(data)
    if not success: logger.error(f"Failed to update DB for new asset: {asset_id_scraped}"); messagebox.showerror("Database Error", f"Failed to save asset {asset_id_scraped} to DB.", parent=parent_dialog); parent_dialog.status_var.set(f"DB error saving {asset_id_scraped}."); return

    try:
        tech_name = os.getenv('USERNAME', 'AuditTool'); site = config.get('site', 'Unknown')
        db.record_scan(asset_id_scraped, "in", tech_name, "Initial check-in (via Audit/ServiceNow lookup)", site)
        logger.info(f"Auto checked in new asset: {asset_id_scraped} at site: {site}")
        messagebox.showinfo("Asset Added", f"Asset {asset_id_scraped} found in ServiceNow, added to database, and checked in to site {site}.", parent=parent_dialog)
        parent_dialog.status_var.set(f"Asset {asset_id_scraped} added. Re-run audit?")
    except Exception as e:
        logger.error(f"Failed to auto-check-in new asset {asset_id_scraped}: {e}", exc_info=True)
        messagebox.showwarning("Check-in Failed", f"Asset {asset_id_scraped} added to DB, but auto check-in failed.", parent=parent_dialog)
        parent_dialog.status_var.set(f"Asset {asset_id_scraped} added, check-in failed.")

# --- Function to show the Audit Dialog ---
def show_audit_dialog(parent, db_manager, site_id, config): # Added config
    """Creates and shows the modal Audit Dialog."""
    dialog = AuditDialog(parent, db_manager, site_id, config) # Pass config
    dialog.transient(parent) # Make it modal relative to parent
    dialog.grab_set() # Grab focus
    parent.wait_window(dialog) # Wait until dialog is closed

def show_bulk_checkout_dialog(parent, db_manager, site_config):
    """Shows a dialog for checking out multiple assets at once."""
    # Use create_properly_sized_dialog for consistent behavior
    dialog = create_properly_sized_dialog("Bulk Check-Out", min_width=700, min_height=600, parent=parent)
    
    # Main container frame (expandable row 0)
    main_container = ttk.Frame(dialog, padding="10")
    main_container.grid(row=0, column=0, sticky="nsew")
    dialog.rowconfigure(0, weight=1)
    dialog.columnconfigure(0, weight=1)
    
    # Configure main container grid
    main_container.rowconfigure(1, weight=1)  # Make text area expand
    main_container.rowconfigure(4, weight=1)  # Make results area expand
    main_container.columnconfigure(0, weight=1)
    
    # Instructions
    ttk.Label(main_container, text="Scan or enter multiple asset tags or serial numbers (one per line):", 
             font=("", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 5))
    
    # Text area for input
    input_frame = ttk.Frame(main_container)
    input_frame.grid(row=1, column=0, sticky="nsew", pady=5)
    input_frame.rowconfigure(0, weight=1)
    input_frame.columnconfigure(0, weight=1)
    
    input_text = tk.Text(input_frame, wrap="none", width=50, height=10)
    input_text.grid(row=0, column=0, sticky="nsew")
    add_context_menu(input_text)
    
    # Scrollbars
    yscroll = ttk.Scrollbar(input_frame, orient="vertical", command=input_text.yview)
    yscroll.grid(row=0, column=1, sticky="ns")
    xscroll = ttk.Scrollbar(input_frame, orient="horizontal", command=input_text.xview)
    xscroll.grid(row=1, column=0, sticky="ew")
    input_text.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
    
    # Options frame
    options_frame = ttk.LabelFrame(main_container, text="Check-Out Options")
    options_frame.grid(row=2, column=0, sticky="ew", pady=10)
    options_frame.columnconfigure(1, weight=1)
    
    # Technician
    ttk.Label(options_frame, text="Technician:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    tech_entry = ttk.Entry(options_frame, width=30)
    tech_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
    tech_entry.insert(0, os.getenv('USERNAME', ''))
    add_context_menu(tech_entry)
    
    # Notes
    ttk.Label(options_frame, text="Notes:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    notes_entry = ttk.Entry(options_frame, width=30)
    notes_entry.grid(row=1, column=1, sticky="ew", padx=5, pady=5)
    notes_entry.insert(0, "Bulk check-out")
    add_context_menu(notes_entry)
    
    # Status frame
    status_frame = ttk.Frame(main_container)
    status_frame.grid(row=3, column=0, sticky="ew", pady=5)
    
    status_var = tk.StringVar(value="Ready to process")
    ttk.Label(status_frame, textvariable=status_var).pack(side="left")
    
    # Results frame with two columns
    results_frame = ttk.LabelFrame(main_container, text="Results")
    results_frame.grid(row=4, column=0, sticky="nsew", pady=5)
    results_frame.columnconfigure(0, weight=1)
    results_frame.columnconfigure(1, weight=1)
    results_frame.rowconfigure(1, weight=1)
    
    # Initially hide results frame
    results_frame.grid_remove()
    
    # Column headers
    ttk.Label(results_frame, text="Checked Out Assets", font=("", 10, "bold")).grid(row=0, column=0, sticky="w", padx=5, pady=5)
    ttk.Label(results_frame, text="Not in Database", font=("", 10, "bold")).grid(row=0, column=1, sticky="w", padx=5, pady=5)
    
    # Two listboxes for results
    checked_out_frame = ttk.Frame(results_frame)
    checked_out_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
    checked_out_frame.rowconfigure(0, weight=1)
    checked_out_frame.columnconfigure(0, weight=1)
    
    not_in_db_frame = ttk.Frame(results_frame)
    not_in_db_frame.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
    not_in_db_frame.rowconfigure(0, weight=1)
    not_in_db_frame.columnconfigure(0, weight=1)
    
    # Checked out assets listbox
    checked_out_listbox = tk.Listbox(checked_out_frame)
    checked_out_listbox.grid(row=0, column=0, sticky="nsew")
    checked_out_scroll = ttk.Scrollbar(checked_out_frame, orient="vertical", command=checked_out_listbox.yview)
    checked_out_scroll.grid(row=0, column=1, sticky="ns")
    checked_out_listbox.configure(yscrollcommand=checked_out_scroll.set)
    
    # Not in database listbox
    not_in_db_listbox = tk.Listbox(not_in_db_frame)
    not_in_db_listbox.grid(row=0, column=0, sticky="nsew")
    not_in_db_scroll = ttk.Scrollbar(not_in_db_frame, orient="vertical", command=not_in_db_listbox.yview)
    not_in_db_scroll.grid(row=0, column=1, sticky="ns")
    not_in_db_listbox.configure(yscrollcommand=not_in_db_scroll.set)
    
    # Storage for asset data
    checked_out_assets = {}  # Format: {display_text: asset_id}
    not_in_db_assets = []    # List of asset ids/tags not found
    
    # Define action functions - MOVED UP BEFORE REFERENCES
    def on_checked_out_double_click(event):
        """Handle double-click on checked out asset"""
        selection = checked_out_listbox.curselection()
        if not selection:
            return
            
        selected_text = checked_out_listbox.get(selection[0])
        asset_id = checked_out_assets.get(selected_text)
        
        if asset_id:
            from ui.dialogs import show_asset_details
            # Don't use topmost attributes at all
            # Instead, temporarily release grab to allow other windows to get focus
            dialog.grab_release()
            # Show the asset details with dialog as parent
            details_window = show_asset_details(db_manager, asset_id, site_config, parent=parent)

    def check_in_selected():
        """Check in the selected asset"""
        selection = checked_out_listbox.curselection()
        if not selection:
            return
            
        selected_text = checked_out_listbox.get(selection[0])
        asset_id = checked_out_assets.get(selected_text)
        
        if asset_id:
            asset_data = db_manager.get_asset_by_id(asset_id)
            if asset_data:
                from ui.dialogs import show_check_in_out_dialog
                dialog.attributes('-topmost', False)  # Allow new window to be on top
                # Don't withdraw, just make it non-topmost
                show_check_in_out_dialog(db_manager, asset_data, default_status="in", 
                                        callback=None, site_config=site_config, parent=dialog)  # Make dialog the parent
                dialog.attributes('-topmost', True)   # Restore topmost after child closes
                dialog.focus_force()  # Force focus back to our dialog
                
                # Remove from listbox if checked in
                current_status = db_manager.get_asset_current_status(asset_id)
                if current_status and current_status.get('status') == 'in':
                    checked_out_listbox.delete(selection[0])
                    del checked_out_assets[selected_text]

    def on_not_in_db_double_click(event):
        """Handle double-click on not in DB asset"""
        selection = not_in_db_listbox.curselection()
        if not selection:
            return
            
        asset_id = not_in_db_listbox.get(selection[0])
        
        # Check if it's likely an asset tag or serial number
        is_asset = asset_id.upper().startswith('GF-')
        
        # Show ServiceNow retrieval dialog
        from services.servicenow import scrape_servicenow
        # Release grab before starting the process
        dialog.grab_release()
        selection_idx = selection[0]  # Store selection index to use later
        
        # For ServiceNow, we can't just set parent, we need to temporarily
        # hide our dialog completely while the servicenow window is active
        dialog.withdraw()
        
        def background_scrape():
            try:
                asset_data = scrape_servicenow(asset_id, is_asset)
                dialog.after(100, lambda: handle_scrape_result(asset_data, selection_idx))
            except Exception as e:
                dialog.after(100, lambda: handle_scrape_error(str(e), selection_idx))
        
        # Start scraping in background
        import threading
        threading.Thread(target=background_scrape).start()

    def handle_scrape_result(asset_data, selection_idx):
        """Handle the result of ServiceNow scraping"""
        # Show our dialog again
        dialog.deiconify()
        # Regain modal behavior
        dialog.grab_set()
        dialog.focus_force()
        
        if not asset_data:
            messagebox.showinfo("Not Found", 
                            "Asset information could not be retrieved from ServiceNow.", 
                            parent=dialog)
            return
        
        # Ensure we have an asset_tag
        if 'asset_tag' not in asset_data and 'asset_id' in asset_data:
            asset_data['asset_tag'] = asset_data['asset_id']
        
        # Update the database
        success = db_manager.update_asset(asset_data)
        
        if success:
            # Automatically check out the asset
            asset_id = asset_data.get('asset_tag', '')
            tech_name = tech_entry.get().strip() or os.getenv('USERNAME', '')
            notes = notes_entry.get().strip() or "Bulk check-out"
            
            checkout_success = db_manager.record_scan(
                asset_id, "out", tech_name, notes, "Out"
            )
            
            if checkout_success:
                # Add to checked out list
                display_text = f"{asset_id} - Added and checked out"
                checked_out_listbox.insert(tk.END, display_text)
                checked_out_assets[display_text] = asset_id
                
                # Remove from not in DB list
                not_in_db_listbox.delete(listbox_selection)
                
                messagebox.showinfo("Success", 
                                 f"Asset {asset_id} added to database and checked out.", 
                                 parent=dialog)
            else:
                messagebox.showerror("Error", 
                                  f"Asset {asset_id} added to database but checkout failed.", 
                                  parent=dialog)
        else:
            messagebox.showerror("Error", 
                               "Failed to add asset to database.", 
                               parent=dialog)
    
    def handle_scrape_error(error_msg, selection_idx):
        """Handle errors during ServiceNow scraping"""
        # Show our dialog again
        dialog.deiconify()
        # Regain modal behavior
        dialog.grab_set()
        dialog.focus_force()
        
        messagebox.showerror("Error", 
                        f"Error retrieving asset information: {error_msg}", 
                        parent=dialog)
        
    # Define context menus for listboxes
    checked_out_menu = tk.Menu(checked_out_listbox, tearoff=0)
    checked_out_menu.add_command(label="Asset Details", 
                                command=lambda: on_checked_out_double_click(None))
    checked_out_menu.add_command(label="Check In", 
                                command=lambda: check_in_selected())
    
    not_in_db_menu = tk.Menu(not_in_db_listbox, tearoff=0)
    not_in_db_menu.add_command(label="Add to Database", 
                              command=lambda: on_not_in_db_double_click(None))
    
    # Show context menu on right-click
    def show_checked_out_menu(event):
        listbox = event.widget
        index = listbox.nearest(event.y)
        if index >= 0:
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(index)
            checked_out_menu.post(event.x_root, event.y_root)
    
    def show_not_in_db_menu(event):
        listbox = event.widget
        index = listbox.nearest(event.y)
        if index >= 0:
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(index)
            not_in_db_menu.post(event.x_root, event.y_root)
    
    # Process function
    def process_checkout():
        # Get inputs
        asset_list_raw = input_text.get("1.0", "end").strip()
        tech_name = tech_entry.get().strip()
        notes = notes_entry.get().strip()
        
        if not asset_list_raw:
            messagebox.showwarning("No Input", "Please enter at least one asset to check out.", parent=dialog)
            return
            
        if not tech_name:
            tech_name = os.getenv('USERNAME', 'Unknown')
        
        # Parse asset list (one per line)
        asset_ids = [line.strip() for line in asset_list_raw.split('\n') if line.strip()]
        
        if not asset_ids:
            messagebox.showwarning("Invalid Input", "Could not parse any valid asset IDs.", parent=dialog)
            return
        
        # Clear previous results
        checked_out_listbox.delete(0, tk.END)
        not_in_db_listbox.delete(0, tk.END)
        checked_out_assets.clear()
        not_in_db_assets.clear()
        
        # Show results frame
        results_frame.grid()
        
        # Process the assets
        success_count = 0
        not_found_count = 0
        already_out_count = 0
        error_count = 0
        
        # Update status
        status_var.set(f"Processing {len(asset_ids)} assets...")
        dialog.update_idletasks()
        
        for asset_id in asset_ids:
            # Try to normalize asset ID
            if asset_id.lower().startswith('gf-'):
                asset_id = asset_id.upper()
                
            # Find asset
            asset_data = db_manager.get_asset_by_id(asset_id)
            
            # If not found by ID, try by serial number
            if not asset_data:
                # Try different case variations for serial numbers
                asset_data = db_manager.get_asset_by_serial(asset_id)
                if not asset_data:
                    asset_data = db_manager.get_asset_by_serial(asset_id.upper())
                if not asset_data:
                    asset_data = db_manager.get_asset_by_serial(asset_id.lower())
            
            # Handle result
            if not asset_data:
                not_found_count += 1
                not_in_db_listbox.insert(tk.END, asset_id)
                not_in_db_assets.append(asset_id)
                continue
            
            # Get actual asset ID to display
            actual_asset_id = asset_data['asset_id']
            
            # Check current status
            current_status = db_manager.get_asset_current_status(actual_asset_id)
            current_state = current_status.get('status', 'unknown') if current_status else 'unknown'
            
            if current_state == 'out':
                already_out_count += 1
                display_text = f"{actual_asset_id} - Already checked out"
                checked_out_listbox.insert(tk.END, display_text)
                checked_out_assets[display_text] = actual_asset_id
                continue
            
            # Record the check-out
            success = db_manager.record_scan(
                actual_asset_id, 
                "out", 
                tech_name, 
                notes or "Bulk check-out", 
                "Out"  # Site is always "Out" for check-outs
            )
            
            if success:
                success_count += 1
                checked_out_listbox.insert(tk.END, actual_asset_id)
                checked_out_assets[actual_asset_id] = actual_asset_id
            else:
                error_count += 1
        
        # Update status
        status_var.set(f"Completed: {success_count} checked out, {not_found_count} not found, {already_out_count} already out, {error_count} errors")
    
    # Add bindings AFTER function definitions
    checked_out_listbox.bind("<Double-1>", on_checked_out_double_click)
    not_in_db_listbox.bind("<Double-1>", on_not_in_db_double_click)
    checked_out_listbox.bind("<Button-3>", show_checked_out_menu)
    not_in_db_listbox.bind("<Button-3>", show_not_in_db_menu)

    # Button frame
    button_frame = ttk.Frame(dialog)
    button_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
    
    process_btn = ttk.Button(button_frame, text="Process Check-Out", command=process_checkout)
    process_btn.pack(side="left", padx=5)
    
    cancel_btn = ttk.Button(button_frame, text="Close", command=dialog.destroy)
    cancel_btn.pack(side="right", padx=5)
    
    # Focus on text input area
    input_text.focus_set()
    
    # Make the dialog modal
    dialog.transient(parent)
    dialog.grab_set()
    parent.wait_window(dialog)

def show_bulk_checkin_dialog(parent, db_manager, site_config):
    """Shows a dialog for checking in multiple assets at once."""
    # Use create_properly_sized_dialog for consistent behavior
    dialog = create_properly_sized_dialog("Bulk Check-In", min_width=700, min_height=600, parent=parent)
    
    # Get the current site from config
    current_site = site_config.get('site', 'Unknown')
    
    # Main container frame (expandable row 0)
    main_container = ttk.Frame(dialog, padding="10")
    main_container.grid(row=0, column=0, sticky="nsew")
    dialog.rowconfigure(0, weight=1)
    dialog.columnconfigure(0, weight=1)
    
    # Configure main container grid
    main_container.rowconfigure(1, weight=1)  # Make text area expand
    main_container.rowconfigure(4, weight=1)  # Make results area expand
    main_container.columnconfigure(0, weight=1)
    
    # Instructions
    ttk.Label(main_container, text="Scan or enter multiple asset tags or serial numbers (one per line):", 
             font=("", 10, "bold")).grid(row=0, column=0, sticky="w", pady=(0, 5))
    
    # Text area for input
    input_frame = ttk.Frame(main_container)
    input_frame.grid(row=1, column=0, sticky="nsew", pady=5)
    input_frame.rowconfigure(0, weight=1)
    input_frame.columnconfigure(0, weight=1)
    
    input_text = tk.Text(input_frame, wrap="none", width=50, height=10)
    input_text.grid(row=0, column=0, sticky="nsew")
    add_context_menu(input_text)
    
    # Scrollbars
    yscroll = ttk.Scrollbar(input_frame, orient="vertical", command=input_text.yview)
    yscroll.grid(row=0, column=1, sticky="ns")
    xscroll = ttk.Scrollbar(input_frame, orient="horizontal", command=input_text.xview)
    xscroll.grid(row=1, column=0, sticky="ew")
    input_text.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)
    
    # Options frame
    options_frame = ttk.LabelFrame(main_container, text="Check-In Options")
    options_frame.grid(row=2, column=0, sticky="ew", pady=10)
    options_frame.columnconfigure(1, weight=1)
    
    # Technician
    ttk.Label(options_frame, text="Technician:").grid(row=0, column=0, sticky="w", padx=5, pady=5)
    tech_entry = ttk.Entry(options_frame, width=30)
    tech_entry.grid(row=0, column=1, sticky="ew", padx=5, pady=5)
    tech_entry.insert(0, os.getenv('USERNAME', ''))
    add_context_menu(tech_entry)
    
    # Site - showing the current site (read-only)
    ttk.Label(options_frame, text="Site:").grid(row=1, column=0, sticky="w", padx=5, pady=5)
    site_label = ttk.Label(options_frame, text=current_site)
    site_label.grid(row=1, column=1, sticky="w", padx=5, pady=5)
    
    # Notes
    ttk.Label(options_frame, text="Notes:").grid(row=2, column=0, sticky="w", padx=5, pady=5)
    notes_entry = ttk.Entry(options_frame, width=30)
    notes_entry.grid(row=2, column=1, sticky="ew", padx=5, pady=5)
    notes_entry.insert(0, "Bulk check-in")
    add_context_menu(notes_entry)
    
    # Status frame
    status_frame = ttk.Frame(main_container)
    status_frame.grid(row=3, column=0, sticky="ew", pady=5)
    
    status_var = tk.StringVar(value="Ready to process")
    ttk.Label(status_frame, textvariable=status_var).pack(side="left")
    
    # Results frame with two columns
    results_frame = ttk.LabelFrame(main_container, text="Results")
    results_frame.grid(row=4, column=0, sticky="nsew", pady=5)
    results_frame.columnconfigure(0, weight=1)
    results_frame.columnconfigure(1, weight=1)
    results_frame.rowconfigure(1, weight=1)
    
    # Initially hide results frame
    results_frame.grid_remove()
    
    # Column headers
    ttk.Label(results_frame, text="Checked In Assets", font=("", 10, "bold")).grid(row=0, column=0, sticky="w", padx=5, pady=5)
    ttk.Label(results_frame, text="Not in Database", font=("", 10, "bold")).grid(row=0, column=1, sticky="w", padx=5, pady=5)
    
    # Two listboxes for results
    checked_in_frame = ttk.Frame(results_frame)
    checked_in_frame.grid(row=1, column=0, sticky="nsew", padx=5, pady=5)
    checked_in_frame.rowconfigure(0, weight=1)
    checked_in_frame.columnconfigure(0, weight=1)
    
    not_in_db_frame = ttk.Frame(results_frame)
    not_in_db_frame.grid(row=1, column=1, sticky="nsew", padx=5, pady=5)
    not_in_db_frame.rowconfigure(0, weight=1)
    not_in_db_frame.columnconfigure(0, weight=1)
    
    # Checked in assets listbox
    checked_in_listbox = tk.Listbox(checked_in_frame)
    checked_in_listbox.grid(row=0, column=0, sticky="nsew")
    checked_in_scroll = ttk.Scrollbar(checked_in_frame, orient="vertical", command=checked_in_listbox.yview)
    checked_in_scroll.grid(row=0, column=1, sticky="ns")
    checked_in_listbox.configure(yscrollcommand=checked_in_scroll.set)
    
    # Not in database listbox
    not_in_db_listbox = tk.Listbox(not_in_db_frame)
    not_in_db_listbox.grid(row=0, column=0, sticky="nsew")
    not_in_db_scroll = ttk.Scrollbar(not_in_db_frame, orient="vertical", command=not_in_db_listbox.yview)
    not_in_db_scroll.grid(row=0, column=1, sticky="ns")
    not_in_db_listbox.configure(yscrollcommand=not_in_db_scroll.set)
    
    # Storage for asset data
    checked_in_assets = {}  # Format: {display_text: asset_id}
    not_in_db_assets = []    # List of asset ids/tags not found
    
    # Define action functions - defined before references
    def on_checked_in_double_click(event):
        """Handle double-click on checked in asset"""
        selection = checked_in_listbox.curselection()
        if not selection:
            return
            
        selected_text = checked_in_listbox.get(selection[0])
        asset_id = checked_in_assets.get(selected_text)
        
        if asset_id:
            from ui.dialogs import show_asset_details
            # Release grab before showing details
            dialog.grab_release()
            show_asset_details(db_manager, asset_id, site_config, parent=parent)
    
    def check_out_selected():
        """Check out the selected asset"""
        selection = checked_in_listbox.curselection()
        if not selection:
            return
            
        selected_text = checked_in_listbox.get(selection[0])
        asset_id = checked_in_assets.get(selected_text)
        
        if asset_id:
            asset_data = db_manager.get_asset_by_id(asset_id)
            if asset_data:
                from ui.dialogs import show_check_in_out_dialog
                # Release grab before showing dialog
                dialog.grab_release()
                show_check_in_out_dialog(db_manager, asset_data, default_status="out", 
                                        callback=None, site_config=site_config, parent=parent)
                
                # Remove from listbox if checked out
                current_status = db_manager.get_asset_current_status(asset_id)
                if current_status and current_status.get('status') == 'out':
                    checked_in_listbox.delete(selection[0])
                    del checked_in_assets[selected_text]
    
    def on_not_in_db_double_click(event):
        """Handle double-click on not in DB asset"""
        selection = not_in_db_listbox.curselection()
        if not selection:
            return
            
        asset_id = not_in_db_listbox.get(selection[0])
        
        # Check if it's likely an asset tag or serial number
        is_asset = asset_id.upper().startswith('GF-')
        
        # Show ServiceNow retrieval dialog
        from services.servicenow import scrape_servicenow
        # Hide our dialog while showing ServiceNow
        dialog.grab_release()
        dialog.withdraw()
        selection_idx = selection[0]  # Store selection index to use later
        
        def background_scrape():
            try:
                asset_data = scrape_servicenow(asset_id, is_asset)
                dialog.after(100, lambda: handle_scrape_result(asset_data, selection_idx))
            except Exception as e:
                dialog.after(100, lambda: handle_scrape_error(str(e), selection_idx))
        
        # Start scraping in background
        import threading
        threading.Thread(target=background_scrape).start()
    
    def handle_scrape_result(asset_data, listbox_selection):
        """Handle the result of ServiceNow scraping"""
        # Show our dialog again
        dialog.deiconify()
        dialog.grab_set()
        dialog.focus_force()
        
        if not asset_data:
            messagebox.showinfo("Not Found", 
                             "Asset information could not be retrieved from ServiceNow.", 
                             parent=dialog)
            return
        
        # Ensure we have an asset_tag
        if 'asset_tag' not in asset_data and 'asset_id' in asset_data:
            asset_data['asset_tag'] = asset_data['asset_id']
        
        # Update the database
        success = db_manager.update_asset(asset_data)
        
        if success:
            # Automatically check in the asset
            asset_id = asset_data.get('asset_tag', '')
            tech_name = tech_entry.get().strip() or os.getenv('USERNAME', '')
            notes = notes_entry.get().strip() or "Bulk check-in"
            
            checkin_success = db_manager.record_scan(
                asset_id, "in", tech_name, notes, current_site
            )
            
            if checkin_success:
                # Add to checked in list
                display_text = f"{asset_id} - Added and checked in"
                checked_in_listbox.insert(tk.END, display_text)
                checked_in_assets[display_text] = asset_id
                
                # Remove from not in DB list
                not_in_db_listbox.delete(listbox_selection)
                
                messagebox.showinfo("Success", 
                                 f"Asset {asset_id} added to database and checked in to {current_site}.", 
                                 parent=dialog)
            else:
                messagebox.showerror("Error", 
                                  f"Asset {asset_id} added to database but check-in failed.", 
                                  parent=dialog)
        else:
            messagebox.showerror("Error", 
                               "Failed to add asset to database.", 
                               parent=dialog)
    
    def handle_scrape_error(error_msg, listbox_selection):
        """Handle errors during ServiceNow scraping"""
        # Show our dialog again
        dialog.deiconify()
        dialog.grab_set()
        dialog.focus_force()
        
        messagebox.showerror("Error", 
                          f"Error retrieving asset information: {error_msg}", 
                          parent=dialog)
    
    # Define context menus for listboxes
    checked_in_menu = tk.Menu(checked_in_listbox, tearoff=0)
    checked_in_menu.add_command(label="Asset Details", 
                                command=lambda: on_checked_in_double_click(None))
    checked_in_menu.add_command(label="Check Out", 
                                command=lambda: check_out_selected())
    
    not_in_db_menu = tk.Menu(not_in_db_listbox, tearoff=0)
    not_in_db_menu.add_command(label="Add to Database", 
                              command=lambda: on_not_in_db_double_click(None))
    
    # Show context menu on right-click
    def show_checked_in_menu(event):
        listbox = event.widget
        index = listbox.nearest(event.y)
        if index >= 0:
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(index)
            checked_in_menu.post(event.x_root, event.y_root)
    
    def show_not_in_db_menu(event):
        listbox = event.widget
        index = listbox.nearest(event.y)
        if index >= 0:
            listbox.selection_clear(0, tk.END)
            listbox.selection_set(index)
            not_in_db_menu.post(event.x_root, event.y_root)
    
    # Process function
    def process_checkin():
        # Get inputs
        asset_list_raw = input_text.get("1.0", "end").strip()
        tech_name = tech_entry.get().strip()
        notes = notes_entry.get().strip()
        
        if not asset_list_raw:
            messagebox.showwarning("No Input", "Please enter at least one asset to check in.", parent=dialog)
            return
            
        if not tech_name:
            tech_name = os.getenv('USERNAME', 'Unknown')
        
        # Parse asset list (one per line)
        asset_ids = [line.strip() for line in asset_list_raw.split('\n') if line.strip()]
        
        if not asset_ids:
            messagebox.showwarning("Invalid Input", "Could not parse any valid asset IDs.", parent=dialog)
            return
        
        # Clear previous results
        checked_in_listbox.delete(0, tk.END)
        not_in_db_listbox.delete(0, tk.END)
        checked_in_assets.clear()
        not_in_db_assets.clear()
        
        # Show results frame
        results_frame.grid()
        
        # Process the assets
        success_count = 0
        not_found_count = 0
        already_in_count = 0
        error_count = 0
        
        # Update status
        status_var.set(f"Processing {len(asset_ids)} assets...")
        dialog.update_idletasks()
        
        for asset_id in asset_ids:
            # Try to normalize asset ID
            if asset_id.lower().startswith('gf-'):
                asset_id = asset_id.upper()
                
            # Find asset
            asset_data = db_manager.get_asset_by_id(asset_id)
            
            # If not found by ID, try by serial number
            if not asset_data:
                # Try different case variations for serial numbers
                asset_data = db_manager.get_asset_by_serial(asset_id)
                if not asset_data:
                    asset_data = db_manager.get_asset_by_serial(asset_id.upper())
                if not asset_data:
                    asset_data = db_manager.get_asset_by_serial(asset_id.lower())
            
            # Handle result
            if not asset_data:
                not_found_count += 1
                not_in_db_listbox.insert(tk.END, asset_id)
                not_in_db_assets.append(asset_id)
                continue
            
            # Get actual asset ID to display
            actual_asset_id = asset_data['asset_id']
            
            # Check current status
            current_status = db_manager.get_asset_current_status(actual_asset_id)
            current_state = current_status.get('status', 'unknown') if current_status else 'unknown'
            current_site_in_db = current_status.get('site', '') if current_status else ''
            
            if current_state == 'in' and current_site_in_db == current_site:
                already_in_count += 1
                display_text = f"{actual_asset_id} - Already checked in at {current_site}"
                checked_in_listbox.insert(tk.END, display_text)
                checked_in_assets[display_text] = actual_asset_id
                continue
            
            # Record the check-in
            success = db_manager.record_scan(
                actual_asset_id, 
                "in", 
                tech_name, 
                notes or "Bulk check-in", 
                current_site  # Use the current site from config
            )
            
            if success:
                success_count += 1
                checked_in_listbox.insert(tk.END, actual_asset_id)
                checked_in_assets[actual_asset_id] = actual_asset_id
            else:
                error_count += 1
        
        # Update status
        status_var.set(f"Completed: {success_count} checked in, {not_found_count} not found, {already_in_count} already in, {error_count} errors")
    
    # Add bindings after function definitions
    checked_in_listbox.bind("<Double-1>", on_checked_in_double_click)
    not_in_db_listbox.bind("<Double-1>", on_not_in_db_double_click)
    checked_in_listbox.bind("<Button-3>", show_checked_in_menu)
    not_in_db_listbox.bind("<Button-3>", show_not_in_db_menu)
    
    # Button frame
    button_frame = ttk.Frame(dialog)
    button_frame.grid(row=1, column=0, sticky="ew", padx=10, pady=10)
    
    process_btn = ttk.Button(button_frame, text="Process Check-In", command=process_checkin)
    process_btn.pack(side="left", padx=5)
    
    cancel_btn = ttk.Button(button_frame, text="Close", command=dialog.destroy)
    cancel_btn.pack(side="right", padx=5)
    
    # Focus on text input area
    input_text.focus_set()
    
    # Make the dialog modal
    dialog.transient(parent)
    dialog.grab_set()
    parent.wait_window(dialog)