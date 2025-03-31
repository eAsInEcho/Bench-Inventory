import tkinter as tk
from tkinter import ttk, messagebox
import webbrowser
import os
import json
from datetime import datetime
import threading
import logging
from database import InventoryDatabase
from servicenow import scrape_servicenow, create_properly_sized_dialog, add_context_menu

def add_mousewheel_scrolling(canvas, container_frame):
    """Add mouse wheel scrolling to a canvas with a container frame"""
    def _on_mousewheel(event):
        canvas.yview_scroll(int(-1*(event.delta/120)), "units")
    
    def _on_frame_enter(event):
        canvas.bind_all("<MouseWheel>", _on_mousewheel)
    
    def _on_frame_leave(event):
        canvas.unbind_all("<MouseWheel>")
    
    container_frame.bind("<Enter>", _on_frame_enter)
    container_frame.bind("<Leave>", _on_frame_leave)

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s',
                   filename='inventory_app.log', filemode='a')
logger = logging.getLogger(__name__)

# Add console handler for immediate feedback
console = logging.StreamHandler()
console.setLevel(logging.INFO)
logger.addHandler(console)

class InventoryApp:
    def __init__(self, root):
        self.root = root
        self.root.title("IT Bench Inventory Manager")
        self.root.geometry("1000x600")
        self.db = InventoryDatabase()
        
        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(root)
        self.notebook.pack(fill='both', expand=True)
        
        # Tab 1: Scan/Input Tab
        self.scan_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.scan_tab, text="Scan/Check-in/out")
        
        # Tab 2: Current Inventory
        self.inventory_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.inventory_tab, text="Current Inventory")
        
        # Tab 3: History (30 days)
        self.history_tab = ttk.Frame(self.notebook)
        self.notebook.add(self.history_tab, text="Recent History")
        
        # Initialize all tabs
        self.setup_scan_tab()
        self.setup_inventory_tab()
        self.setup_history_tab()
        
        logger.info("Application initialized")

    def setup_scan_tab(self):
        # Frame for scan input
        input_frame = ttk.LabelFrame(self.scan_tab, text="Asset Input")
        input_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Radio buttons for input type
        self.input_type = tk.StringVar(value="asset")
        ttk.Radiobutton(input_frame, text="Asset Tag", variable=self.input_type, 
                        value="asset").grid(row=0, column=0, padx=5, pady=5)
        ttk.Radiobutton(input_frame, text="Serial Number", variable=self.input_type, 
                        value="serial").grid(row=0, column=1, padx=5, pady=5)
        
        # Entry field for scanning/typing
        ttk.Label(input_frame, text="Scan or Enter:").grid(row=1, column=0, padx=5, pady=5)
        self.input_entry = ttk.Entry(input_frame, width=30)
        self.input_entry.grid(row=1, column=1, padx=5, pady=5)
        self.input_entry.focus()  # Auto-focus for scanner
        
        # Add context menu for input entry
        add_context_menu(self.input_entry)
        
        # Bind Enter key to processing the input
        self.input_entry.bind("<Return>", lambda e: self.process_asset_lookup())
        
        # Action buttons
        button_frame = ttk.Frame(input_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Look Up Asset", width=15,
                  command=self.process_asset_lookup).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Manual Entry", width=15,
                  command=self.show_manual_entry_form).pack(side="left", padx=5)
        
        # Results display
        ttk.Label(input_frame, text="Asset Details:").grid(row=3, column=0, padx=5, pady=5, sticky="nw")
        self.details_text = tk.Text(input_frame, width=60, height=15, wrap="word")
        self.details_text.grid(row=3, column=1, padx=5, pady=5)
        self.details_text.config(state="disabled")  # Read-only initially
        
        # Add context menu for details
        add_context_menu(self.details_text)
        
        # Status bar for showing scraping progress
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(input_frame, textvariable=self.status_var).grid(row=4, column=0, columnspan=2, pady=5)

        # More radio button code...
        ttk.Radiobutton(input_frame, text="Serial Number", variable=self.input_type, 
                        value="serial").grid(row=0, column=1, padx=5, pady=5)
        
        # Entry field for scanning/typing
        ttk.Label(input_frame, text="Scan or Enter:").grid(row=1, column=0, padx=5, pady=5)
        self.input_entry = ttk.Entry(input_frame, width=30)
        self.input_entry.grid(row=1, column=1, padx=5, pady=5)
        self.input_entry.focus()  # Auto-focus for scanner
        
        # Add context menu for input entry
        add_context_menu(self.input_entry)
        
        # Bind Enter key to processing the input
        self.input_entry.bind("<Return>", lambda e: self.process_asset_lookup())
        
        # Action buttons
        button_frame = ttk.Frame(input_frame)
        button_frame.grid(row=2, column=0, columnspan=2, pady=10)
        
        ttk.Button(button_frame, text="Look Up Asset", width=15,
                  command=self.process_asset_lookup).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Manual Entry", width=15,
                  command=self.show_manual_entry_form).pack(side="left", padx=5)
        
        # Results display
        ttk.Label(input_frame, text="Asset Details:").grid(row=3, column=0, padx=5, pady=5, sticky="nw")
        self.details_text = tk.Text(input_frame, width=60, height=15, wrap="word")
        self.details_text.grid(row=3, column=1, padx=5, pady=5)
        self.details_text.config(state="disabled")  # Read-only initially
        
        # Add context menu for details
        add_context_menu(self.details_text)
        
        # Status bar for showing scraping progress
        self.status_var = tk.StringVar(value="Ready")
        ttk.Label(input_frame, textvariable=self.status_var).grid(row=4, column=0, columnspan=2, pady=5)

    def setup_inventory_tab(self):
        # Frame for inventory
        inventory_frame = ttk.Frame(self.inventory_tab)
        inventory_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Controls
        control_frame = ttk.Frame(inventory_frame)
        control_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(control_frame, text="Refresh", 
                  command=self.refresh_inventory).pack(side="left", padx=5)
        ttk.Button(control_frame, text="Export", 
                  command=self.export_inventory).pack(side="left", padx=5)
                  
        # Add toggle for showing deleted assets
        self.show_deleted = tk.BooleanVar(value=False)
        ttk.Checkbutton(control_frame, text="Show Deleted Assets", 
                       variable=self.show_deleted, 
                       command=self.refresh_inventory).pack(side="left", padx=20)
        
        # Search frame
        search_frame = ttk.Frame(inventory_frame)
        search_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Label(search_frame, text="Search:").pack(side="left", padx=5)
        self.inventory_search = ttk.Entry(search_frame, width=30)
        self.inventory_search.pack(side="left", padx=5)
        self.inventory_search.bind("<KeyRelease>", self.filter_inventory)
        
        # Add context menu
        add_context_menu(self.inventory_search)
        
        # Treeview for inventory list
        columns = ("asset_id", "serial", "make_model", "status", "owner", "check_in_date")
        self.inventory_tree = ttk.Treeview(inventory_frame, columns=columns, show="headings")
        
        # Define headings
        self.inventory_tree.heading("asset_id", text="Asset Tag", command=lambda: self.treeview_sort_column(self.inventory_tree, "asset_id", False))
        self.inventory_tree.heading("serial", text="Serial Number", command=lambda: self.treeview_sort_column(self.inventory_tree, "serial", False))
        self.inventory_tree.heading("make_model", text="Make/Model", command=lambda: self.treeview_sort_column(self.inventory_tree, "make_model", False))
        self.inventory_tree.heading("status", text="Status", command=lambda: self.treeview_sort_column(self.inventory_tree, "status", False))
        self.inventory_tree.heading("owner", text="Assigned To", command=lambda: self.treeview_sort_column(self.inventory_tree, "owner", False))
        self.inventory_tree.heading("check_in_date", text="Check-in Date", command=lambda: self.treeview_sort_column(self.inventory_tree, "check_in_date", False))
        
        # Define columns
        self.inventory_tree.column("asset_id", width=100)
        self.inventory_tree.column("serial", width=150)
        self.inventory_tree.column("make_model", width=200)
        self.inventory_tree.column("status", width=100)
        self.inventory_tree.column("owner", width=150)
        self.inventory_tree.column("check_in_date", width=150)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(inventory_frame, orient="vertical", command=self.inventory_tree.yview)
        self.inventory_tree.configure(yscrollcommand=scrollbar.set)
        
        self.inventory_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind double-click to view details
        self.inventory_tree.bind("<Double-1>", self.on_inventory_double_click)
        
        # Add context menu to treeview
        self.add_inventory_context_menu()
        
        # Populate with current data
        self.refresh_inventory()

    def setup_history_tab(self):
        # Frame for history
        history_frame = ttk.Frame(self.history_tab)
        history_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Controls
        control_frame = ttk.Frame(history_frame)
        control_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(control_frame, text="Refresh", 
                  command=self.refresh_history).pack(side="left", padx=5)
        ttk.Button(control_frame, text="Search Older Records", 
                  command=self.search_full_history).pack(side="left", padx=5)
        
        # Treeview for history list
        columns = ("timestamp", "asset_id", "serial", "status", "tech", "notes")
        self.history_tree = ttk.Treeview(history_frame, columns=columns, show="headings")
        
        # Define headings with sorting
        self.history_tree.heading("timestamp", text="Date/Time", command=lambda: self.treeview_sort_column(self.history_tree, "timestamp", False))
        self.history_tree.heading("asset_id", text="Asset Tag", command=lambda: self.treeview_sort_column(self.history_tree, "asset_id", False))
        self.history_tree.heading("serial", text="Serial Number", command=lambda: self.treeview_sort_column(self.history_tree, "serial", False))
        self.history_tree.heading("status", text="Action", command=lambda: self.treeview_sort_column(self.history_tree, "status", False))
        self.history_tree.heading("tech", text="Technician", command=lambda: self.treeview_sort_column(self.history_tree, "tech", False))
        self.history_tree.heading("notes", text="Notes", command=lambda: self.treeview_sort_column(self.history_tree, "notes", False))
        
        # Define columns
        self.history_tree.column("timestamp", width=150)
        self.history_tree.column("asset_id", width=100)
        self.history_tree.column("serial", width=150)
        self.history_tree.column("status", width=80)
        self.history_tree.column("tech", width=100)
        self.history_tree.column("notes", width=250)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(history_frame, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=scrollbar.set)
        
        self.history_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind double-click to view details
        self.history_tree.bind("<Double-1>", self.on_history_double_click)
        
        # Populate with recent data
        self.refresh_history()

    def treeview_sort_column(self, treeview, col, reverse):
        """Sort treeview content when column header is clicked"""
        l = [(treeview.set(k, col), k) for k in treeview.get_children('')]
        
        # Try to sort numerically if possible, otherwise sort as strings
        try:
            l.sort(key=lambda t: float(t[0]), reverse=reverse)
        except ValueError:
            l.sort(reverse=reverse)
        
        # Rearrange items in sorted positions
        for index, (val, k) in enumerate(l):
            treeview.move(k, '', index)
        
        # Reverse sort next time
        treeview.heading(col, command=lambda: self.treeview_sort_column(treeview, col, not reverse))

    def add_inventory_context_menu(self):
        """Add context menu to inventory treeview"""
        menu = tk.Menu(self.inventory_tree, tearoff=0)
        
        def popup(event):
            # Select row under mouse
            iid = self.inventory_tree.identify_row(event.y)
            if iid:
                self.inventory_tree.selection_set(iid)
                menu.post(event.x_root, event.y_root)
        
        menu.add_command(label="View Details", command=self.view_selected_asset)
        menu.add_command(label="Check In", command=lambda: self.check_inout_selected_asset("in"))
        menu.add_command(label="Check Out", command=lambda: self.check_inout_selected_asset("out"))
        menu.add_separator()
        menu.add_command(label="Edit Asset", command=self.edit_selected_asset)
        menu.add_command(label="Delete Asset", command=self.delete_selected_asset)
        
        self.inventory_tree.bind("<Button-3>", popup)

    def view_selected_asset(self):
        """Show details for the selected asset"""
        selection = self.inventory_tree.selection()
        if not selection:
            return
        
        asset_id = self.inventory_tree.item(selection[0], "values")[0]
        self.show_asset_details(asset_id)
    
    def check_inout_selected_asset(self, action):
        """Check in or out the selected asset"""
        selection = self.inventory_tree.selection()
        if not selection:
            return
        
        asset_id = self.inventory_tree.item(selection[0], "values")[0]
        asset_data = self.db.get_asset_by_id(asset_id)
        
        if asset_data:
            self.show_check_in_out_dialog(asset_data, default_status=action)
    
    def edit_selected_asset(self):
        """Edit the selected asset"""
        selection = self.inventory_tree.selection()
        if not selection:
            return
        
        asset_id = self.inventory_tree.item(selection[0], "values")[0]
        self.edit_asset(asset_id)
    
    def delete_selected_asset(self):
        """Delete the selected asset"""
        selection = self.inventory_tree.selection()
        if not selection:
            return
        
        asset_id = self.inventory_tree.item(selection[0], "values")[0]
        self.delete_asset(asset_id)
    
    def on_inventory_double_click(self, event):
        """Handle double-click on inventory item"""
        item = self.inventory_tree.identify('item', event.x, event.y)
        if item:
            asset_id = self.inventory_tree.item(item, "values")[0]
            self.show_asset_details(asset_id)
    
    def on_history_double_click(self, event):
        """Handle double-click on history item"""
        item = self.history_tree.identify('item', event.x, event.y)
        if item:
            asset_id = self.history_tree.item(item, "values")[1]
            self.show_asset_details(asset_id)

    def show_asset_details(self, asset_id):
        """Show detailed information about an asset"""
        asset = self.db.get_asset_by_id(asset_id)
        if not asset:
            messagebox.showerror("Error", f"Asset {asset_id} not found")
            return
        
        # Get recent history
        history = self.db.get_asset_history(asset_id, 5)
        
        # Create detail window
        detail_window = create_properly_sized_dialog(f"Asset Details: {asset_id}", 800, 600)
        
        # Create tabs
        notebook = ttk.Notebook(detail_window)
        notebook.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Tab 1: Asset Details
        details_tab = ttk.Frame(notebook)
        notebook.add(details_tab, text="Asset Details")
        
        # Tab 2: History
        history_tab = ttk.Frame(notebook)
        notebook.add(history_tab, text="Recent History")
        
        # Create scrollable frame for details
        details_frame = ttk.Frame(details_tab)
        details_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        canvas = tk.Canvas(details_frame)
        scrollbar = ttk.Scrollbar(details_frame, orient="vertical", command=canvas.yview)
        
        scrollable_frame = ttk.Frame(canvas)
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Add mouse wheel scrolling
        add_mousewheel_scrolling(canvas, details_frame)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Populate details tab
        row = 0
        for key in asset.keys():
            if key in ['asset_id', 'asset_tag']:
                field_name = "Asset Tag"
            elif key == 'cmdb_url':
                field_name = "CMDB URL"
            else:
                # Convert snake_case to Title Case
                field_name = " ".join(word.capitalize() for word in key.split('_'))
            
            ttk.Label(scrollable_frame, text=f"{field_name}:").grid(row=row, column=0, sticky="w", padx=10, pady=5)
            
            # Use Text widget for values to allow copying
            value_text = tk.Text(scrollable_frame, width=40, height=1, wrap="word")
            value_text.grid(row=row, column=1, sticky="ew", padx=10, pady=5)
            value_text.insert("1.0", str(asset[key] or ""))
            value_text.config(state="disabled")
            add_context_menu(value_text)
            
            row += 1
        
        # Add CMDB URL button if available
        def open_cmdb():
            if asset.get('cmdb_url'):
                webbrowser.open(asset['cmdb_url'])
            else:
                # Generate a URL based on asset tag
                url = f"https://globalfoundries.service-now.com/u_cmdb_ci_notebook.do?sysparm_query=asset_tag%3D{asset_id}"
                webbrowser.open(url)
        
        ttk.Button(scrollable_frame, text="View in CMDB", command=open_cmdb).grid(
            row=row, column=0, columnspan=2, pady=10)
        
        # Add action buttons
        row += 1
        action_frame = ttk.Frame(scrollable_frame)
        action_frame.grid(row=row, column=0, columnspan=2, pady=10)
        
        ttk.Button(action_frame, text="Check In", 
                  command=lambda: self.show_check_in_out_dialog(asset, "in")).pack(side="left", padx=5)
        ttk.Button(action_frame, text="Check Out", 
                  command=lambda: self.show_check_in_out_dialog(asset, "out")).pack(side="left", padx=5)
        ttk.Button(action_frame, text="Edit", 
                  command=lambda: self.edit_asset(asset_id)).pack(side="left", padx=5)
        ttk.Button(action_frame, text="Delete", 
                  command=lambda: self.delete_asset(asset_id, detail_window)).pack(side="left", padx=5)
        
        # Populate history tab
        history_frame = ttk.Frame(history_tab)
        history_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Create history treeview
        columns = ("timestamp", "status", "tech", "notes")
        history_tree = ttk.Treeview(history_frame, columns=columns, show="headings")
        
        history_tree.heading("timestamp", text="Date/Time")
        history_tree.heading("status", text="Action")
        history_tree.heading("tech", text="Technician")
        history_tree.heading("notes", text="Notes")
        
        history_tree.column("timestamp", width=150)
        history_tree.column("status", width=80)
        history_tree.column("tech", width=100)
        history_tree.column("notes", width=250)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(history_frame, orient="vertical", command=history_tree.yview)
        history_tree.configure(yscrollcommand=scrollbar.set)
        
        history_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Add context menu for history
        add_context_menu(history_tree)
        
        # Populate history
        for item in history:
            history_tree.insert("", "end", values=(
                item['timestamp'],
                item['status'],
                item['tech_name'],
                item['notes']
            ))

    def process_asset_lookup(self):
        """Process asset lookup with check for existing assets"""
        identifier = self.input_entry.get().strip()
        if not identifier:
            messagebox.showerror("Error", "Please enter an asset tag or serial number")
            return
        
        is_asset = self.input_type.get() == "asset"
        
        # Update status
        self.status_var.set(f"Looking up {'asset' if is_asset else 'serial'}: {identifier}...")
        self.root.update()
        
        # First check if it exists in the database
        existing_asset = None
        if is_asset:
            existing_asset = self.db.get_asset_by_id(identifier)
        else:
            existing_asset = self.db.get_asset_by_serial(identifier)
        
        if existing_asset:
            # Check if this is a deleted asset
            is_deleted = existing_asset.get('operational_status') == 'DELETED'
            
            # Asset exists - show options
            title = "Asset Found" if not is_deleted else "Deleted Asset Found"
            option_dialog = create_properly_sized_dialog(title, 500, 300)
            
            if not is_deleted:
                ttk.Label(option_dialog, text=f"Asset already exists in the system:").pack(pady=10)
            else:
                ttk.Label(option_dialog, text=f"This asset was previously deleted from inventory:", font=("", 10, "bold")).pack(pady=10)
                ttk.Label(option_dialog, text=f"You can reactivate it or start fresh with data from CMDB.").pack(pady=5)
            
            ttk.Label(option_dialog, text=f"Asset Tag: {existing_asset['asset_id']}").pack(anchor="w", padx=20)
            ttk.Label(option_dialog, text=f"Serial: {existing_asset.get('serial_number', '')}").pack(anchor="w", padx=20)
            ttk.Label(option_dialog, text=f"Model: {existing_asset.get('manufacturer', '')} {existing_asset.get('model_description', '')}").pack(anchor="w", padx=20)
            
            button_frame = ttk.Frame(option_dialog)
            button_frame.pack(pady=20)
            
            def view_details():
                option_dialog.destroy()
                self.show_asset_details(existing_asset['asset_id'])
            
            def check_in():
                option_dialog.destroy()
                self.show_check_in_out_dialog(existing_asset, "in")
            
            def check_out():
                option_dialog.destroy()
                self.show_check_in_out_dialog(existing_asset, "out") 
                
            def view_cmdb():
                if existing_asset.get('cmdb_url'):
                    import webbrowser
                    webbrowser.open(existing_asset['cmdb_url'])
                else:
                    # Generate a URL based on asset tag or serial
                    asset_id = existing_asset['asset_id']
                    url = f"https://globalfoundries.service-now.com/u_cmdb_ci_notebook.do?sysparm_query=asset_tag%3D{asset_id}"
                    webbrowser.open(url)
            
            def reactivate_asset():
                """Reactivate a previously deleted asset"""
                option_dialog.destroy()
                
                # Update asset status
                asset_data = existing_asset.copy()
                asset_data['operational_status'] = 'Operational'  # Or some other status
                
                # Update database
                self.db.update_asset(asset_data)
                
                # Record reactivation
                self.db.record_scan(asset_data['asset_id'], "reactivated", 
                                   os.getenv('USERNAME', ''), "Asset reactivated from deleted state")
                
                # Show check in dialog
                self.show_check_in_out_dialog(asset_data)
            
            def get_fresh_data():
                """Get fresh data from CMDB instead of using old data"""
                option_dialog.destroy()
                
                identifier = existing_asset['asset_id']
                is_asset = True  # Assuming we always have asset_id
                
                # Run normal scraping procedure as if it's a new asset
                self.status_var.set(f"Scraping data for asset: {identifier}...")
                self.root.update()
                
                # Run scraping in a separate thread
                def scrape_thread():
                    try:
                        asset_data = scrape_servicenow(identifier, is_asset)
                        self.root.after(0, lambda: self.handle_scrape_result(asset_data))
                    except Exception as e:
                        error_message = str(e)
                        logger.error(f"Error in scrape thread: {error_message}")
                        self.root.after(0, lambda msg=error_message: self.status_var.set(f"Error: {msg}"))
                
                threading.Thread(target=scrape_thread, daemon=True).start()
            
            # Show different buttons based on whether asset is deleted
            if not is_deleted:
                ttk.Button(button_frame, text="View Details", command=view_details).pack(side="left", padx=5)
                ttk.Button(button_frame, text="View in CMDB", command=view_cmdb).pack(side="left", padx=5)
            
            # Clear input and status
            self.input_entry.delete(0, "end")
            self.status_var.set("Ready")
            
        else:
            # Asset doesn't exist - proceed with normal workflow
            # Update status
            self.status_var.set(f"Scraping data for {'asset' if is_asset else 'serial'}: {identifier}...")
            self.root.update()
            
            # Run scraping in a separate thread to keep UI responsive
            def scrape_thread():
                try:
                    asset_data = scrape_servicenow(identifier, is_asset)
                    
                    # Update UI in the main thread
                    self.root.after(0, lambda: self.handle_scrape_result(asset_data))
                    
                except Exception as e:
                    error_message = str(e)
                    logger.error(f"Error in scrape thread: {error_message}")
                    self.root.after(0, lambda msg=error_message: self.status_var.set(f"Error: {msg}"))
            
            # Start the thread
            threading.Thread(target=scrape_thread, daemon=True).start()

    def handle_scrape_result(self, asset_data):
        """Handle the result of scraping"""
        if not asset_data:
            self.status_var.set("Failed to retrieve asset data")
            messagebox.showerror("Error", "Could not retrieve asset data from ServiceNow")
            return
        
        # Update the database
        self.db.update_asset(asset_data)
        
        # Display asset details
        self.display_asset_details(asset_data)
        
        # Update status
        self.status_var.set("Ready")
        
        # Show check-in/out dialog
        self.show_check_in_out_dialog(asset_data)
    
    def display_asset_details(self, asset_data):
        """Display asset details in the main window"""
        self.details_text.config(state="normal")
        self.details_text.delete("1.0", "end")
        
        details = f"""Asset Tag: {asset_data.get('asset_tag', '')}
Serial Number: {asset_data.get('serial_number', '')}
Hostname: {asset_data.get('hostname', '')}
Make/Model: {asset_data.get('manufacturer', '')} {asset_data.get('model_description', '')}
Assigned To: {asset_data.get('assigned_to', '')}
Location: {asset_data.get('location', '')}
OS: {asset_data.get('os', '')} {asset_data.get('os_version', '')}
Warranty Expires: {asset_data.get('warranty_expiration', '')}

Comments: {asset_data.get('comments', '')}
"""
        
        self.details_text.insert("1.0", details)
        self.details_text.config(state="disabled")
    
    def show_check_in_out_dialog(self, asset_data, default_status="in"):
        """Show dialog to check asset in or out"""
        check_window = create_properly_sized_dialog("Check In/Out Asset", 400, 350)
        
        # Display asset info
        info_frame = ttk.LabelFrame(check_window, text="Asset Information")
        info_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Label(info_frame, text=f"Asset Tag: {asset_data['asset_id'] if 'asset_id' in asset_data else asset_data['asset_tag']}").pack(anchor="w", padx=10, pady=2)
        ttk.Label(info_frame, text=f"Model: {asset_data.get('manufacturer', '')} {asset_data.get('model_description', '')}").pack(anchor="w", padx=10, pady=2)
        ttk.Label(info_frame, text=f"S/N: {asset_data.get('serial_number', '')}").pack(anchor="w", padx=10, pady=2)
        
        # Action selection
        action_frame = ttk.LabelFrame(check_window, text="Action")
        action_frame.pack(fill="x", padx=10, pady=10)
        
        status_var = tk.StringVar(value=default_status)
        ttk.Radiobutton(action_frame, text="Check IN to Bench", variable=status_var, value="in").pack(anchor="w", padx=10, pady=5)
        ttk.Radiobutton(action_frame, text="Check OUT from Bench", variable=status_var, value="out").pack(anchor="w", padx=10, pady=5)
        
        # Notes
        notes_frame = ttk.LabelFrame(check_window, text="Notes")
        notes_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        ttk.Label(notes_frame, text="Technician:").pack(anchor="w", padx=10, pady=2)
        tech_entry = ttk.Entry(notes_frame, width=30)
        tech_entry.pack(fill="x", padx=10, pady=2)
        tech_entry.insert(0, os.getenv('USERNAME', ''))
        
        # Add context menu
        add_context_menu(tech_entry)
        
        ttk.Label(notes_frame, text="Notes:").pack(anchor="w", padx=10, pady=2)
        notes_text = tk.Text(notes_frame, width=40, height=3)
        notes_text.pack(fill="both", expand=True, padx=10, pady=2)
        
        # Add context menu
        add_context_menu(notes_text)
        
        # Completion function
        def complete_action():
            asset_id = asset_data['asset_id'] if 'asset_id' in asset_data else asset_data['asset_tag']
            status = status_var.get()
            tech = tech_entry.get().strip() or os.getenv('USERNAME', '')
            notes = notes_text.get("1.0", "end-1c")
            
            # Record in database
            self.db.record_scan(asset_id, status, tech, notes)
            
            # Update UI
            self.refresh_inventory()
            self.refresh_history()
            
            # Clear input
            self.input_entry.delete(0, "end")
            self.input_entry.focus()
            
            # Close window
            check_window.destroy()
            
            messagebox.showinfo("Success", f"Asset {asset_id} checked {status}")
        
        # Buttons
        button_frame = ttk.Frame(check_window)
        button_frame.pack(fill="x", padx=10, pady=10)
        
        ttk.Button(button_frame, text="Complete", command=complete_action).pack(side="left", padx=5)
        ttk.Button(button_frame, text="Cancel", command=check_window.destroy).pack(side="right", padx=5)

    def edit_asset(self, asset_id):
        """Edit an existing asset"""
        asset = self.db.get_asset_by_id(asset_id)
        if not asset:
            messagebox.showerror("Error", f"Asset {asset_id} not found")
            return
        
        # Create edit form
        entry_window = create_properly_sized_dialog("Edit Asset", 700, 800)
        
        # Instructions
        ttk.Label(entry_window, text=f"Edit asset {asset_id}:").pack(pady=10)
        
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
        add_mousewheel_scrolling(canvas, main_frame)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Field entry widgets
        fields = [
            "hostname", "serial_number", "operational_status",
            "install_status", "location", "ci_region", "owned_by", 
            "assigned_to", "manufacturer", "model_id", "model_description",
            "vendor", "warranty_expiration", "os", "os_version"
        ]
        
        field_labels = {
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
        
        # Add asset tag (readonly)
        ttk.Label(scrollable_frame, text="Asset Tag:").grid(row=0, column=0, sticky="w", padx=10, pady=5)
        asset_tag_entry = ttk.Entry(scrollable_frame, width=50)
        asset_tag_entry.grid(row=0, column=1, padx=10, pady=5)
        asset_tag_entry.insert(0, asset_id)
        asset_tag_entry.config(state="readonly")
        
        # Add other fields
        for i, field in enumerate(fields, 1):
            ttk.Label(scrollable_frame, text=f"{field_labels[field]}:").grid(row=i, column=0, sticky="w", padx=10, pady=5)
            entry = ttk.Entry(scrollable_frame, width=50)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entries[field] = entry
            
            # Pre-fill with existing data
            if asset.get(field):
                entry.insert(0, asset[field])
            
            # Add context menu
            add_context_menu(entry)
        
        # Comments field (multiline)
        ttk.Label(scrollable_frame, text="Comments:").grid(row=len(fields)+1, column=0, sticky="nw", padx=10, pady=5)
        comments_text = tk.Text(scrollable_frame, width=50, height=5)
        comments_text.grid(row=len(fields)+1, column=1, padx=10, pady=5)
        if asset.get('comments'):
            comments_text.insert("1.0", asset['comments'])
        
        # Add context menu
        add_context_menu(comments_text)
        
        # CMDB URL field
        ttk.Label(scrollable_frame, text="CMDB URL:").grid(row=len(fields)+2, column=0, sticky="w", padx=10, pady=5)
        cmdb_url_entry = ttk.Entry(scrollable_frame, width=50)
        cmdb_url_entry.grid(row=len(fields)+2, column=1, padx=10, pady=5)
        if asset.get('cmdb_url'):
            cmdb_url_entry.insert(0, asset['cmdb_url'])
        
        # Add context menu
        add_context_menu(cmdb_url_entry)
        
        # CMDB View function
        def view_cmdb():
            if asset.get('cmdb_url'):
                webbrowser.open(asset['cmdb_url'])
            else:
                # Generate a URL based on asset tag
                url = f"https://globalfoundries.service-now.com/u_cmdb_ci_notebook.do?sysparm_query=asset_tag%3D{asset_id}"
                webbrowser.open(url)
                
        # Save function
        def save_data():
            # Collect data from entries
            asset_data = {field: entries[field].get().strip() for field in fields}
            asset_data['comments'] = comments_text.get("1.0", "end-1c")
            asset_data['cmdb_url'] = cmdb_url_entry.get().strip()
            asset_data['asset_tag'] = asset_id
            
            # Update database
            success = self.db.update_asset(asset_data)
            if not success:
                messagebox.showerror("Error", "Failed to update asset")
                return
            
            # Record edit in history
            self.db.record_scan(asset_id, "edited", os.getenv('USERNAME', ''), "Asset details edited")
            
            # Refresh UI
            self.refresh_inventory()
            self.refresh_history()
            
            # Close window
            entry_window.destroy()
            
            messagebox.showinfo("Success", f"Asset {asset_id} updated")
        
        # Buttons
        button_frame = ttk.Frame(entry_window)
        button_frame.pack(fill="x", padx=10, pady=20)
        
        ttk.Button(button_frame, text="Save", command=save_data).pack(side="left", padx=10)
        ttk.Button(button_frame, text="View in CMDB", command=view_cmdb).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Cancel", command=entry_window.destroy).pack(side="right", padx=10)

    def show_manual_entry_form(self, identifier=None):
        """Show manual data entry form"""
        
        entry_window = create_properly_sized_dialog("Manual Asset Entry", 700, 800)
        
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
        add_mousewheel_scrolling(canvas, main_frame)
        
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
            "asset_tag": "Asset Tag",
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
            entry = ttk.Entry(scrollable_frame, width=50)
            entry.grid(row=i, column=1, padx=10, pady=5)
            entries[field] = entry
            
            # Add context menu
            add_context_menu(entry)
            
            # Pre-fill the identifier
            if identifier and ((field == "asset_tag" and identifier.startswith("GF-")) or 
                              (field == "serial_number" and not identifier.startswith("GF-"))):
                entry.insert(0, identifier)
        
        # Comments field (multiline)
        ttk.Label(scrollable_frame, text="Comments:").grid(row=len(fields), column=0, sticky="nw", padx=10, pady=5)
        comments_text = tk.Text(scrollable_frame, width=50, height=5)
        comments_text.grid(row=len(fields), column=1, padx=10, pady=5)
        
        # Add context menu
        add_context_menu(comments_text)
        
        # CMDB URL field
        ttk.Label(scrollable_frame, text="CMDB URL:").grid(row=len(fields)+1, column=0, sticky="w", padx=10, pady=5)
        cmdb_url_entry = ttk.Entry(scrollable_frame, width=50)
        cmdb_url_entry.grid(row=len(fields)+1, column=1, padx=10, pady=5)
        
        # Add context menu
        add_context_menu(cmdb_url_entry)
        
        # Save function
        def collect_data():
            asset_data = {field: entries[field].get() for field in fields}
            asset_data['comments'] = comments_text.get("1.0", "end-1c")
            asset_data['cmdb_url'] = cmdb_url_entry.get()
            
            if not asset_data['asset_tag']:
                messagebox.showerror("Error", "Asset Tag is required")
                return None
            
            return asset_data
        
        def save_data():
            asset_data = collect_data()
            if not asset_data:
                return
            
            # Update database
            success = self.db.update_asset(asset_data)
            if not success:
                messagebox.showerror("Error", "Failed to save asset data")
                return
                
            # Close window
            entry_window.destroy()
            
            # Show check-in/out dialog
            self.show_check_in_out_dialog(asset_data)
        
        # Buttons
        button_frame = ttk.Frame(scrollable_frame)
        button_frame.grid(row=len(fields)+2, column=0, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="Save", command=save_data).pack(side="left", padx=10)
        ttk.Button(button_frame, text="Cancel", command=entry_window.destroy).pack(side="left", padx=10)

    def refresh_inventory(self):
        """Refresh the current inventory view"""
        # Clear existing items
        for item in self.inventory_tree.get_children():
            self.inventory_tree.delete(item)
        
        # Get current inventory, including deleted items if toggle is set
        include_deleted = self.show_deleted.get() if hasattr(self, 'show_deleted') else False
        inventory = self.db.get_current_inventory(include_deleted)
        
        # Populate treeview
        for item in inventory:
            make_model = f"{item.get('manufacturer', '')} {item.get('model_description', '')}"
            
            # Determine if this is a deleted asset
            is_deleted = item.get('operational_status') == 'DELETED'
            
            # Insert with tags for deleted assets (for styling)
            tags = ('deleted',) if is_deleted else ()
            
            self.inventory_tree.insert("", "end", values=(
                item.get('asset_id', ''),
                item.get('serial_number', ''),
                make_model.strip(),
                item.get('operational_status', ''),
                item.get('assigned_to', ''),
                item.get('check_in_date', '')
            ), tags=tags)
        
        # Configure tag for deleted items (gray and italic)
        self.inventory_tree.tag_configure('deleted', foreground='gray', font=('', 9, 'italic'))
    
    def refresh_history(self):
        """Refresh the recent history view"""
        # Clear existing items
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        
        # Get recent history
        history = self.db.get_recent_history()
        
        # Populate treeview
        for item in history:
            self.history_tree.insert("", "end", values=(
                item.get('timestamp', ''),
                item.get('asset_id', ''),
                item.get('serial_number', ''),
                item.get('status', ''),
                item.get('tech_name', ''),
                item.get('notes', '')
            ))
    
    def filter_inventory(self, event=None):
        """Filter inventory based on search text"""
        search_text = self.inventory_search.get().lower()
        
        # Clear existing items
        for item in self.inventory_tree.get_children():
            self.inventory_tree.delete(item)
        
        # Get all inventory
        include_deleted = self.show_deleted.get() if hasattr(self, 'show_deleted') else False
        inventory = self.db.get_current_inventory(include_deleted)
        
        # Filter and populate
        for item in inventory:
            if (search_text in str(item.get('asset_id', '')).lower() or
                search_text in str(item.get('serial_number', '')).lower() or
                search_text in str(item.get('manufacturer', '')).lower() or
                search_text in str(item.get('model_description', '')).lower() or
                search_text in str(item.get('assigned_to', '')).lower()):
                
                make_model = f"{item.get('manufacturer', '')} {item.get('model_description', '')}"
                
                # Determine if this is a deleted asset
                is_deleted = item.get('operational_status') == 'DELETED'
                
                # Insert with tags for deleted assets (for styling)
                tags = ('deleted',) if is_deleted else ()
                
                self.inventory_tree.insert("", "end", values=(
                    item.get('asset_id', ''),
                    item.get('serial_number', ''),
                    make_model.strip(),
                    item.get('operational_status', ''),
                    item.get('assigned_to', ''),
                    item.get('check_in_date', '')
                ), tags=tags)

    def search_full_history(self):
        """Open dialog to search complete history"""
        search_window = create_properly_sized_dialog("Search Full History", 400, 150)
        
        ttk.Label(search_window, text="Enter Asset Tag or Serial Number:").pack(pady=10)
        
        search_entry = ttk.Entry(search_window, width=30)
        search_entry.pack(pady=5)
        search_entry.focus()
        
        # Add context menu
        add_context_menu(search_entry)
        
        def perform_search():
            search_term = search_entry.get().strip()
            if not search_term:
                return
                
            # Search database
            results = self.db.search_asset_history(search_term)
            
            if not results:
                messagebox.showinfo("No Results", "No history found for this search term")
                return
                
            # Create results window
            results_window = create_properly_sized_dialog(f"History for {search_term}", 800, 400)
            
            # Create treeview for results
            columns = ("timestamp", "asset_id", "status", "tech", "notes")
            results_tree = ttk.Treeview(results_window, columns=columns, show="headings")
            
            # Add sorting to columns
            results_tree.heading("timestamp", text="Date/Time", 
                                command=lambda: self.treeview_sort_column(results_tree, "timestamp", False))
            results_tree.heading("asset_id", text="Asset Tag", 
                                command=lambda: self.treeview_sort_column(results_tree, "asset_id", False))
            results_tree.heading("status", text="Action", 
                                command=lambda: self.treeview_sort_column(results_tree, "status", False))
            results_tree.heading("tech", text="Technician", 
                                command=lambda: self.treeview_sort_column(results_tree, "tech", False))
            results_tree.heading("notes", text="Notes", 
                                command=lambda: self.treeview_sort_column(results_tree, "notes", False))
            
            results_tree.column("timestamp", width=150)
            results_tree.column("asset_id", width=100)
            results_tree.column("status", width=80)
            results_tree.column("tech", width=100)
            results_tree.column("notes", width=300)
            
            scrollbar = ttk.Scrollbar(results_window, orient="vertical", command=results_tree.yview)
            results_tree.configure(yscrollcommand=scrollbar.set)
            
            results_tree.pack(side="left", fill="both", expand=True)
            scrollbar.pack(side="right", fill="y")
            
            # Add context menu to results
            add_context_menu(results_tree)
            
            # Double-click to view asset details
            def on_double_click(event):
                item = results_tree.identify('item', event.x, event.y)
                if item:
                    asset_id = results_tree.item(item, "values")[1]
                    self.show_asset_details(asset_id)
            
            results_tree.bind("<Double-1>", on_double_click)
            
            # Populate results
            for item in results:
                results_tree.insert("", "end", values=(
                    item.get('timestamp', ''),
                    item.get('asset_id', ''),
                    item.get('status', ''),
                    item.get('tech_name', ''),
                    item.get('notes', '')
                ))
        
        ttk.Button(search_window, text="Search", command=perform_search).pack(pady=10)
        
        # Bind Enter key to search
        search_entry.bind("<Return>", lambda e: perform_search())
    
    def export_inventory(self):
        """Export current inventory to CSV"""
        # Ask if deleted assets should be included
        include_deleted = False
        
        # If show_deleted setting exists, use it as default
        if hasattr(self, 'show_deleted'):
            include_deleted = self.show_deleted.get()
        
        # Ask user to confirm export settings
        result = messagebox.askyesno("Export Options", 
                                   f"Include deleted assets in export?\n\nCurrent view setting: {'Include' if include_deleted else 'Exclude'} deleted assets",
                                   icon='question')
        
        # Override based on user choice
        include_deleted = result
        
        filename = f"inventory_export_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        
        inventory = self.db.get_current_inventory(include_deleted)
        
        if not inventory:
            messagebox.showinfo("No Data", "No inventory data to export")
            return
        
        with open(filename, 'w') as f:
            # Write headers
            headers = list(inventory[0].keys())
            f.write(','.join([f'"{h}"' for h in headers]) + '\n')
            
            # Write data
            for item in inventory:
                row = [f'"{str(item.get(h, ""))}"' for h in headers]
                f.write(','.join(row) + '\n')
        
        asset_count = len(inventory)
        deleted_count = sum(1 for item in inventory if item.get('operational_status') == 'DELETED')
        
        messagebox.showinfo("Export Complete", 
                           f"Inventory exported to {filename}\n\nTotal assets: {asset_count}\nDeleted assets: {deleted_count}")
    
    def delete_asset(self, asset_id, parent_window=None):
        """Delete an asset from inventory"""
        if messagebox.askyesno("Confirm Delete", 
                              f"Are you sure you want to delete asset {asset_id}?\n\nThis will mark the asset as deleted and remove it from the regular inventory view but keep its details and history.",
                              parent=parent_window):
            
            # Record the deletion in history first
            self.db.record_scan(asset_id, "deleted", os.getenv('USERNAME', ''), "Asset deleted from inventory")
            
            # Soft delete the asset (mark as deleted)
            success = self.db.delete_asset(asset_id)
            
            if success:
                messagebox.showinfo("Success", f"Asset {asset_id} deleted from inventory. If scanned again, you'll have the option to restore it.", parent=parent_window)
                
                # Refresh UI
                self.refresh_inventory()
                self.refresh_history()
                
                # Close parent window if provided
                if parent_window:
                    parent_window.destroy()
            else:
                messagebox.showerror("Error", f"Failed to delete asset {asset_id}", parent=parent_window)

if __name__ == "__main__":
    root = tk.Tk()
    app = InventoryApp(root)
    root.mainloop()
