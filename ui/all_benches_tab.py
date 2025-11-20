import tkinter as tk
from tkinter import ttk, messagebox
import os
import logging

from ui.utils import add_context_menu
from ui.dialogs import show_asset_details, create_properly_sized_dialog
from ui.dialogs import format_timestamp

logger = logging.getLogger(__name__)

class AllBenchesTab:
    def __init__(self, parent, db, app, config=None):
        self.db = db
        self.app = app
        self.config = config or {}
        self.frame = ttk.Frame(parent)
        self.setup_ui()
        
    def setup_ui(self):
        """Set up the All Benches tab UI"""
        # Frame for inventory
        inventory_frame = ttk.Frame(self.frame)
        inventory_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Controls
        control_frame = ttk.Frame(inventory_frame)
        control_frame.pack(fill="x", padx=5, pady=5)
        
        # Refresh and Export buttons
        ttk.Button(control_frame, text="Refresh", 
                  command=self.refresh_inventory).pack(side="left", padx=5)
        ttk.Button(control_frame, text="Export", 
                  command=self.export_inventory).pack(side="left", padx=5)
                  
        # Add toggle for showing deleted assets
        self.show_deleted = tk.BooleanVar(value=False)
        ttk.Checkbutton(control_frame, text="Show Deleted Assets", 
                       variable=self.show_deleted, 
                       command=self.refresh_inventory).pack(side="left", padx=20)
        
        # Site filter dropdown
        site_filter_frame = ttk.Frame(control_frame)
        site_filter_frame.pack(side="right", padx=20)
        
        ttk.Label(site_filter_frame, text="Filter by Site:").pack(side="left", padx=5)
        self.site_filter = ttk.Combobox(site_filter_frame, width=15, state="readonly")
        self.site_filter.pack(side="left", padx=5)
        self.site_filter.bind("<<ComboboxSelected>>", self.refresh_inventory)
        
        # Populate site dropdown
        self.populate_site_dropdown()
        
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
        columns = ("asset_id", "hostname", "serial", "model", "assigned_to", "check_in_date", "site")
        self.inventory_tree = ttk.Treeview(inventory_frame, columns=columns, show="headings")
        
        # Define headings
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
    
    def populate_site_dropdown(self):
        """Populate the site dropdown with available sites"""
        # Get all unique sites from the database
        conn = None
        try:
            conn = self.db.get_connection(write=False)  # Read operation
            cursor = self.db.cursor(conn)
            
            # Query to get unique sites
            if self.db.using_local:
                cursor.execute("SELECT DISTINCT site FROM scan_history WHERE site IS NOT NULL AND site != ''")
            else:
                cursor.execute("SELECT DISTINCT site FROM scan_history WHERE site IS NOT NULL AND site != ''")
            
            sites = [row[0] for row in cursor.fetchall()]
            
            # Add "All Sites" option
            sites = ["All Sites"] + sorted(sites)
            
            # Update the dropdown
            self.site_filter['values'] = sites
            
            # Set default value to "All Sites"
            self.site_filter.current(0)
            
        except Exception as e:
            logger.error(f"Error populating site dropdown: {str(e)}")
        finally:
            if conn:
                self.db.release_connection(conn)
    
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
        
        self.inventory_tree.bind("<Button-3>", popup)

    def view_selected_asset(self):
        """Show details for the selected asset"""
        selection = self.inventory_tree.selection()
        if not selection:
            return
        
        asset_id = self.inventory_tree.item(selection[0], "values")[0]
        show_asset_details(self.db, asset_id, self.config)
    
    def on_inventory_double_click(self, event):
        """Handle double-click on inventory item"""
        item = self.inventory_tree.identify('item', event.x, event.y)
        if item:
            asset_id = self.inventory_tree.item(item, "values")[0]
            show_asset_details(self.db, asset_id, self.config)

    def refresh_inventory(self, event=None):
        """Refresh the inventory view"""
        # Clear existing items
        for item in self.inventory_tree.get_children():
            self.inventory_tree.delete(item)
        
        # Get current inventory, including deleted items if toggle is set
        include_deleted = self.show_deleted.get() if hasattr(self, 'show_deleted') else False
        selected_site = self.site_filter.get() if hasattr(self, 'site_filter') else "All Sites"
        
        # Get inventory from database
        inventory = self.db.get_current_inventory(include_deleted)
        
        # Filter by site if needed
        if selected_site and selected_site != "All Sites":
            filtered_inventory = []
            for item in inventory:
                site = item.get('site')
                if site == selected_site:
                    filtered_inventory.append(item)
            inventory = filtered_inventory
        
        # Populate treeview
        for item in inventory:
            # Determine if this is a deleted asset
            is_deleted = item.get('operational_status') == 'DELETED'
            
            # Insert with tags for deleted assets (for styling)
            tags = ('deleted',) if is_deleted else ()
            
            # Add site highlighting for current site
            current_site = self.config.get('site')
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
        selected_site = self.site_filter.get() if hasattr(self, 'site_filter') else "All Sites"
        
        # Get inventory from database
        inventory = self.db.get_current_inventory(include_deleted)
        
        # Filter by site if needed
        if selected_site and selected_site != "All Sites":
            filtered_inventory = []
            for item in inventory:
                site = item.get('site')
                if site == selected_site:
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
                
                # Add site highlighting for current site
                current_site = self.config.get('site')
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
        """Export the filtered inventory to CSV"""
        from datetime import datetime
        
        # Get configurations for export
        include_deleted = self.show_deleted.get() if hasattr(self, 'show_deleted') else False
        selected_site = self.site_filter.get() if hasattr(self, 'site_filter') else "All Sites"
        
        # Ask user to confirm export settings
        export_options = f"Include deleted assets: {'Yes' if include_deleted else 'No'}\n"
        export_options += f"Site filter: {selected_site}"
        
        result = messagebox.askyesno("Export Options", 
                                   f"Confirm export with these settings?\n\n{export_options}",
                                   icon='question')
        
        if not result:
            return
        
        # Get inventory from database
        inventory = self.db.get_current_inventory(include_deleted)
        
        # Filter by site if needed
        if selected_site and selected_site != "All Sites":
            filtered_inventory = []
            for item in inventory:
                site = item.get('site')
                if site == selected_site:
                    filtered_inventory.append(item)
            inventory = filtered_inventory
        
        if not inventory:
            messagebox.showinfo("No Data", "No inventory data to export")
            return
        
        # Generate filename with site information
        site_text = selected_site.replace(" ", "_").lower()
        filename = f"all_benches_inventory_{site_text}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        
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
        site_count = len(set(item.get('site', '') for item in inventory))
        
        messagebox.showinfo("Export Complete", 
                           f"Inventory exported to {filename}\n\n"
                           f"Total assets: {asset_count}\n"
                           f"Deleted assets: {deleted_count}\n"
                           f"Sites included: {site_count}")