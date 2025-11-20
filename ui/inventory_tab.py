import tkinter as tk
from tkinter import ttk, messagebox
import os
import logging

from ui.dialogs import show_asset_details, create_properly_sized_dialog, format_timestamp, show_audit_dialog
from ui.utils import add_context_menu
from ui.dialogs import format_timestamp

logger = logging.getLogger(__name__)

class InventoryTab:
    def __init__(self, parent, db, app, config=None):
        self.db = db
        self.app = app
        self.config = config or {}
        self.frame = ttk.Frame(parent)
        self.setup_ui()

    def setup_ui(self):
        """Set up the inventory tab UI"""
        # Frame for inventory
        inventory_frame = ttk.Frame(self.frame)
        inventory_frame.pack(fill="both", expand=True, padx=10, pady=10)

        # --- Controls ---
        control_frame = ttk.Frame(inventory_frame)
        control_frame.pack(fill="x", padx=5, pady=5)

        ttk.Button(control_frame, text="Refresh",
                command=self.refresh_inventory).pack(side="left", padx=5)
        ttk.Button(control_frame, text="Export",
                command=self.export_inventory).pack(side="left", padx=5)
        ttk.Button(control_frame, text="Audit Bench",
                command=self.open_audit_dialog).pack(side="left", padx=5)

        # Add toggle for showing deleted assets
        self.show_deleted = tk.BooleanVar(value=False)
        ttk.Checkbutton(control_frame, text="Show Deleted Assets",
                        variable=self.show_deleted,
                        command=self.refresh_inventory).pack(side="left", padx=20)

        # Site indicator
        site = self.config.get('site', 'Unknown')
        site_label = ttk.Label(control_frame, text=f"Site: {site}", font=("", 10, "bold"))
        site_label.pack(side="right", padx=20)

        # Search frame
        search_frame = ttk.Frame(inventory_frame)
        search_frame.pack(fill="x", padx=5, pady=5)

        ttk.Label(search_frame, text="Search:").pack(side="left", padx=5)
        self.inventory_search = ttk.Entry(search_frame, width=30)
        self.inventory_search.pack(side="left", padx=5)
        self.inventory_search.bind("<KeyRelease>", self.filter_inventory)

        # Add context menu
        add_context_menu(self.inventory_search)

        # Updated columns for treeview
        columns = ("asset_id", "hostname", "serial", "model", "assigned_to", "check_in_date", "site")
        self.inventory_tree = ttk.Treeview(inventory_frame, columns=columns, show="headings")

        # Define headings with updated labels
        self.inventory_tree.heading("asset_id", text="Asset Tag", command=lambda: self.treeview_sort_column(self.inventory_tree, "asset_id", False))
        self.inventory_tree.heading("hostname", text="Hostname", command=lambda: self.treeview_sort_column(self.inventory_tree, "hostname", False))
        self.inventory_tree.heading("serial", text="Serial", command=lambda: self.treeview_sort_column(self.inventory_tree, "serial", False))
        self.inventory_tree.heading("model", text="Model", command=lambda: self.treeview_sort_column(self.inventory_tree, "model", False))
        self.inventory_tree.heading("assigned_to", text="Assigned To", command=lambda: self.treeview_sort_column(self.inventory_tree, "assigned_to", False))
        self.inventory_tree.heading("check_in_date", text="Check-In Date", command=lambda: self.treeview_sort_column(self.inventory_tree, "check_in_date", False))
        self.inventory_tree.heading("site", text="Site", command=lambda: self.treeview_sort_column(self.inventory_tree, "site", False))

        # Define columns widths
        self.inventory_tree.column("asset_id", width=100)
        self.inventory_tree.column("hostname", width=150)
        self.inventory_tree.column("serial", width=150)
        self.inventory_tree.column("model", width=200)
        self.inventory_tree.column("assigned_to", width=150)
        self.inventory_tree.column("check_in_date", width=150)
        self.inventory_tree.column("site", width=80)

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
        menu.add_command(label="Check In", command=self.check_in_selected_asset)
        menu.add_separator()
        menu.add_command(label="Edit Asset", command=self.edit_selected_asset)
        
        self.inventory_tree.bind("<Button-3>", popup)

    def view_selected_asset(self):
        """Show details for the selected asset"""
        selection = self.inventory_tree.selection()
        if not selection:
            return
        
        asset_id = self.inventory_tree.item(selection[0], "values")[0]
        show_asset_details(self.db, asset_id, self.config)  # Pass config here
    
    def check_in_selected_asset(self):
        """Check in the selected asset"""
        from ui.dialogs import show_check_in_out_dialog
        
        selection = self.inventory_tree.selection()
        if not selection:
            return
        
        asset_id = self.inventory_tree.item(selection[0], "values")[0]
        asset_data = self.db.get_asset_by_id(asset_id)
        
        if asset_data:
            show_check_in_out_dialog(self.db, asset_data, default_status="in", callback=self.app.refresh_all, site_config=self.config)
    
    def edit_selected_asset(self):
        """Edit the selected asset"""
        from ui.dialogs import edit_asset
        
        selection = self.inventory_tree.selection()
        if not selection:
            return
        
        asset_id = self.inventory_tree.item(selection[0], "values")[0]
        edit_asset(self.db, asset_id, callback=self.app.refresh_all)
    
    def on_inventory_double_click(self, event):
        """Handle double-click on inventory item"""
        item = self.inventory_tree.identify('item', event.x, event.y)
        if item:
            asset_id = self.inventory_tree.item(item, "values")[0]
            show_asset_details(self.db, asset_id, self.config)  # Pass config here

    def refresh_inventory(self):
        """Refresh the inventory view"""
        # Clear existing items
        for item in self.inventory_tree.get_children():
            self.inventory_tree.delete(item)
        
        # Get current inventory, including deleted items if toggle is set
        include_deleted = self.show_deleted.get() if hasattr(self, 'show_deleted') else False
        current_site = self.config.get('site')
        
        # Get inventory from database
        inventory = self.db.get_current_inventory(include_deleted)
        
        # Filter by site (always filter to current site, since we removed the "Show All Sites" option)
        if current_site:
            filtered_inventory = []
            for item in inventory:
                # Use the site value directly from the database query
                site = item.get('site')
                
                # Include if site matches the current site or site is not specified
                if site == current_site or site is None or site == '' or site == 'Unknown':
                    filtered_inventory.append(item)
            inventory = filtered_inventory
        
        # Populate treeview with updated column data
        for item in inventory:
            # Determine if this is a deleted asset
            is_deleted = item.get('operational_status') == 'DELETED'
            
            # Insert with tags for deleted assets (for styling)
            tags = ('deleted',) if is_deleted else ()
            
            # Add site highlighting
            if item.get('site') == current_site:
                tags = tags + ('current_site',)
            
            # Format model information (combining manufacturer and model)
            model_text = f"{item.get('manufacturer', '')} {item.get('model_id', '')}"
            
            self.inventory_tree.insert("", "end", values=(
                item.get('asset_id', ''),
                item.get('hostname', ''),
                item.get('serial_number', ''),
                model_text.strip(),
                item.get('assigned_to', ''),
                format_timestamp(item.get('check_in_date', '')),
                item.get('site', '')
            ), tags=tags)
        
        # Configure tag for deleted items (gray and italic)
        self.inventory_tree.tag_configure('deleted', foreground='gray', font=('', 9, 'italic'))
        
        # Configure tag for current site items (bold)
        self.inventory_tree.tag_configure('current_site', font=('', 9, 'bold'))
    
    def filter_inventory(self, event=None):
        """Filter inventory based on search text"""
        search_text = self.inventory_search.get().lower()
        
        # Clear existing items
        for item in self.inventory_tree.get_children():
            self.inventory_tree.delete(item)
        
        # Get all inventory
        include_deleted = self.show_deleted.get() if hasattr(self, 'show_deleted') else False
        current_site = self.config.get('site')
        
        # Get current inventory from database - FIX: Changed from get_checked_out_inventory to get_current_inventory
        inventory = self.db.get_current_inventory(include_deleted)
        
        # Filter by site (always filter to current site)
        filtered_inventory = []
        if current_site:
            for item in inventory:
                site = item.get('site')
                if site == current_site or site is None or site == '' or site == 'Unknown':
                    filtered_inventory.append(item)
            inventory = filtered_inventory
        
        # Filter and populate based on search text
        for item in inventory:
            if (search_text in str(item.get('asset_id', '')).lower() or
                search_text in str(item.get('hostname', '')).lower() or
                search_text in str(item.get('serial_number', '')).lower() or
                search_text in str(item.get('manufacturer', '')).lower() or
                search_text in str(item.get('model_id', '')).lower() or
                search_text in str(item.get('model_description', '')).lower() or
                search_text in str(item.get('assigned_to', '')).lower() or
                search_text in str(item.get('site', '')).lower()):
                
                # Determine if this is a deleted asset
                is_deleted = item.get('operational_status') == 'DELETED'
                
                # Insert with tags for deleted assets (for styling)
                tags = ('deleted',) if is_deleted else ()
                
                # Add site highlighting
                if item.get('site') == current_site:
                    tags = tags + ('current_site',)
                
                # Format model information
                model_text = f"{item.get('manufacturer', '')} {item.get('model_id', '')}"
                
                self.inventory_tree.insert("", "end", values=(
                    item.get('asset_id', ''),
                    item.get('hostname', ''),
                    item.get('serial_number', ''),
                    model_text.strip(),
                    item.get('assigned_to', ''),
                    format_timestamp(item.get('check_in_date', '')),
                    item.get('site', '')
                ), tags=tags)
        
        # Configure tag for deleted items and current site items
        self.inventory_tree.tag_configure('deleted', foreground='gray', font=('', 9, 'italic'))
        self.inventory_tree.tag_configure('current_site', font=('', 9, 'bold'))
    
    def export_inventory(self):
        """Export checked out inventory to CSV"""
        from datetime import datetime
        
        # Get configurations for export
        include_deleted = self.show_deleted.get() if hasattr(self, 'show_deleted') else False
        current_site = self.config.get('site')
        
        # Ask user to confirm export settings
        export_options = f"Include deleted assets: {'Yes' if include_deleted else 'No'}\n"
        
        result = messagebox.askyesno("Export Options", 
                                   f"Confirm export with these settings?\n\n{export_options}",
                                   icon='question')
        
        if not result:
            return
            
        # Get inventory based on settings
        inventory = self.db.get_checked_out_inventory(include_deleted)
        
        # Filter by site (always filter to current site)
        if current_site:
            filtered_inventory = []
            for item in inventory:
                # Get the latest scan record for this asset to check its site
                history = self.db.get_asset_history(item['asset_id'], 1)
                site = history[0].get('site') if history else None
                
                # Include if site matches or site is not specified
                if site == current_site or site is None:
                    item['site'] = site or 'Unknown'
                    filtered_inventory.append(item)
            inventory = filtered_inventory
        else:
            # Add site information for all inventory
            for item in inventory:
                history = self.db.get_asset_history(item['asset_id'], 1)
                site = history[0].get('site') if history else None
                item['site'] = site or 'Unknown'
        
        if not inventory:
            messagebox.showinfo("No Data", "No checked out inventory data to export")
            return
        
        # Generate filename with site information
        site_text = current_site if current_site else "unknown_site"
        filename = f"checked_out_inventory_{site_text}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        
        with open(filename, 'w') as f:
            # Write headers (ensure 'site' is included)
            headers = list(inventory[0].keys())
            if 'site' not in headers:
                headers.append('site')
            f.write(','.join([f'"{h}"' for h in headers]) + '\n')
            
            # Write data
            for item in inventory:
                row = [f'"{str(item.get(h, ""))}"' for h in headers]
                f.write(','.join(row) + '\n')
        
        # Report stats
        asset_count = len(inventory)
        deleted_count = sum(1 for item in inventory if item.get('operational_status') == 'DELETED')
        
        messagebox.showinfo("Export Complete", 
                           f"Checked out inventory exported to {filename}\n\n"
                           f"Total assets: {asset_count}\n"
                           f"Deleted assets: {deleted_count}")

    def open_audit_dialog(self):
            """Opens the Bench Audit dialog."""
            current_site = self.config.get('site')
            if not current_site:
                messagebox.showerror("Configuration Error", "Current site is not defined in the configuration.", parent=self.frame)
                return

            # Call the helper function from dialogs.py, passing config
            show_audit_dialog(self.frame, self.db, current_site, self.config) # Pass self.config here