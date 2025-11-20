import tkinter as tk
from tkinter import ttk, messagebox
import os
import logging

from ui.utils import add_context_menu
from ui.dialogs import show_asset_details, create_properly_sized_dialog, format_timestamp

logger = logging.getLogger(__name__)

class FlaggedAssetsTab:
    def __init__(self, parent, db, app, config=None):
        self.db = db
        self.app = app
        self.config = config or {}
        self.frame = ttk.Frame(parent)
        self.setup_ui()
        
    def setup_ui(self):
        """Set up the Flagged Assets tab UI"""
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
        
        # Add title label for clarity
        title_label = ttk.Label(control_frame, text="Flagged Assets", font=("", 12, "bold"))
        title_label.pack(side="right", padx=20)
        
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
        columns = ("asset_id", "hostname", "serial", "model", "flag_timestamp", "flag_tech", "flag_notes", "status", "site")
        self.inventory_tree = ttk.Treeview(inventory_frame, columns=columns, show="headings")
        
        # Define headings
        self.inventory_tree.heading("asset_id", text="Asset Tag", command=lambda: self.treeview_sort_column(self.inventory_tree, "asset_id", False))
        self.inventory_tree.heading("hostname", text="Hostname", command=lambda: self.treeview_sort_column(self.inventory_tree, "hostname", False))
        self.inventory_tree.heading("serial", text="Serial", command=lambda: self.treeview_sort_column(self.inventory_tree, "serial", False))
        self.inventory_tree.heading("model", text="Model", command=lambda: self.treeview_sort_column(self.inventory_tree, "model", False))
        self.inventory_tree.heading("flag_timestamp", text="Flagged Date", command=lambda: self.treeview_sort_column(self.inventory_tree, "flag_timestamp", False))
        self.inventory_tree.heading("flag_tech", text="Flagged By", command=lambda: self.treeview_sort_column(self.inventory_tree, "flag_tech", False))
        self.inventory_tree.heading("flag_notes", text="Flag Reason", command=lambda: self.treeview_sort_column(self.inventory_tree, "flag_notes", False))
        self.inventory_tree.heading("status", text="Status", command=lambda: self.treeview_sort_column(self.inventory_tree, "status", False))
        self.inventory_tree.heading("site", text="Site", command=lambda: self.treeview_sort_column(self.inventory_tree, "site", False))
        
        # Define columns widths
        self.inventory_tree.column("asset_id", width=100)
        self.inventory_tree.column("hostname", width=120)
        self.inventory_tree.column("serial", width=120)
        self.inventory_tree.column("model", width=150)
        self.inventory_tree.column("flag_timestamp", width=120)
        self.inventory_tree.column("flag_tech", width=100)
        self.inventory_tree.column("flag_notes", width=200)
        self.inventory_tree.column("status", width=80)
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
        menu.add_command(label="Unflag Asset", command=self.unflag_selected_asset)
        menu.add_separator()
        menu.add_command(label="Check In", command=self.check_in_selected_asset)
        menu.add_command(label="Check Out", command=self.check_out_selected_asset)
        
        self.inventory_tree.bind("<Button-3>", popup)

    def view_selected_asset(self):
        """Show details for the selected asset"""
        selection = self.inventory_tree.selection()
        if not selection:
            return
        
        asset_id = self.inventory_tree.item(selection[0], "values")[0]
        show_asset_details(self.db, asset_id, self.config)
    
    def unflag_selected_asset(self):
        """Unflag the selected asset"""
        selection = self.inventory_tree.selection()
        if not selection:
            return
        
        asset_id = self.inventory_tree.item(selection[0], "values")[0]
        
        if messagebox.askyesno("Confirm Unflag", 
                              f"Are you sure you want to remove the flag from asset {asset_id}?"):
            # Prompt for reason
            from tkinter import simpledialog
            unflag_reason = simpledialog.askstring("Unflag Reason", 
                                                 "Enter a reason for unflagging this asset:", 
                                                 parent=self.frame)
            
            tech_name = os.getenv('USERNAME', '')
            success = self.db.unflag_asset(asset_id, tech_name, unflag_reason or "Flag removed")
            
            if success:
                messagebox.showinfo("Success", f"Flag removed from asset {asset_id}")
                self.refresh_inventory()
            else:
                messagebox.showerror("Error", f"Failed to remove flag from asset {asset_id}")
    
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
    
    def check_out_selected_asset(self):
        """Check out the selected asset"""
        from ui.dialogs import show_check_in_out_dialog
        
        selection = self.inventory_tree.selection()
        if not selection:
            return
        
        asset_id = self.inventory_tree.item(selection[0], "values")[0]
        asset_data = self.db.get_asset_by_id(asset_id)
        
        if asset_data:
            show_check_in_out_dialog(self.db, asset_data, default_status="out", callback=self.app.refresh_all, site_config=self.config)
    
    def on_inventory_double_click(self, event):
        """Handle double-click on inventory item"""
        item = self.inventory_tree.identify('item', event.x, event.y)
        if item:
            asset_id = self.inventory_tree.item(item, "values")[0]
            show_asset_details(self.db, asset_id, self.config)

    def refresh_inventory(self, event=None):
        """Refresh the flagged inventory view"""
        # Clear existing items
        for item in self.inventory_tree.get_children():
            self.inventory_tree.delete(item)
        
        # Get all flagged assets
        inventory = self.db.get_flagged_assets()
        
        # Populate treeview
        for item in inventory:
            # Get current status/site information
            current_status = item.get('status', 'unknown')
            current_site = item.get('site', '')
            
            # Format model information (combining manufacturer and model)
            model_text = f"{item.get('manufacturer', '')} {item.get('model_id', '')}"
            
            # Format flag timestamp
            flag_timestamp = format_timestamp(item.get('flag_timestamp', ''))
            
            # Format flag notes (truncate if too long)
            flag_notes = item.get('flag_notes', '')
            if len(flag_notes) > 50:
                flag_notes = flag_notes[:47] + "..."
            
            # Set status color tag
            tags = ('flagged',)
            if current_status == 'in':
                tags = ('flagged', 'in_status')
            elif current_status == 'out':
                tags = ('flagged', 'out_status')
            
            self.inventory_tree.insert("", "end", values=(
                item.get('asset_id', ''),
                item.get('hostname', ''),
                item.get('serial_number', ''),
                model_text.strip(),
                flag_timestamp,
                item.get('flag_tech', ''),
                flag_notes,
                current_status.upper() if current_status else '',
                current_site
            ), tags=tags)
        
        # Configure tag styling
        self.inventory_tree.tag_configure('flagged', background='#FFF3E0')  # Light orange background
        self.inventory_tree.tag_configure('in_status', foreground='#388E3C')  # Green text for "IN"
        self.inventory_tree.tag_configure('out_status', foreground='#D32F2F')  # Red text for "OUT"
    
    def filter_inventory(self, event=None):
        """Filter inventory based on search text"""
        search_text = self.inventory_search.get().lower()
        
        # Clear existing items
        for item in self.inventory_tree.get_children():
            self.inventory_tree.delete(item)
        
        # Get all flagged assets
        inventory = self.db.get_flagged_assets()
        
        # Filter and populate based on search text
        for item in inventory:
            if (search_text in str(item.get('asset_id', '')).lower() or
                search_text in str(item.get('hostname', '')).lower() or
                search_text in str(item.get('serial_number', '')).lower() or
                search_text in str(item.get('manufacturer', '')).lower() or
                search_text in str(item.get('model_id', '')).lower() or
                search_text in str(item.get('flag_tech', '')).lower() or
                search_text in str(item.get('flag_notes', '')).lower() or
                search_text in str(item.get('site', '')).lower()):
                
                # Get current status/site information
                current_status = item.get('status', 'unknown')
                current_site = item.get('site', '')
                
                # Format model information
                model_text = f"{item.get('manufacturer', '')} {item.get('model_id', '')}"
                
                # Format flag timestamp
                flag_timestamp = format_timestamp(item.get('flag_timestamp', ''))
                
                # Format flag notes (truncate if too long)
                flag_notes = item.get('flag_notes', '')
                if len(flag_notes) > 50:
                    flag_notes = flag_notes[:47] + "..."
                
                # Set status color tag
                tags = ('flagged',)
                if current_status == 'in':
                    tags = ('flagged', 'in_status')
                elif current_status == 'out':
                    tags = ('flagged', 'out_status')
                
                self.inventory_tree.insert("", "end", values=(
                    item.get('asset_id', ''),
                    item.get('hostname', ''),
                    item.get('serial_number', ''),
                    model_text.strip(),
                    flag_timestamp,
                    item.get('flag_tech', ''),
                    flag_notes,
                    current_status.upper() if current_status else '',
                    current_site
                ), tags=tags)
        
        # Configure tag styling
        self.inventory_tree.tag_configure('flagged', background='#FFF3E0')  # Light orange background
        self.inventory_tree.tag_configure('in_status', foreground='#388E3C')  # Green text for "IN"
        self.inventory_tree.tag_configure('out_status', foreground='#D32F2F')  # Red text for "OUT"
    
    def export_inventory(self):
        """Export the flagged inventory to CSV"""
        from datetime import datetime
        
        # Get all flagged assets
        inventory = self.db.get_flagged_assets()
        
        if not inventory:
            messagebox.showinfo("No Data", "No flagged assets to export")
            return
        
        # Generate filename with timestamp
        filename = f"flagged_assets_{datetime.now().strftime('%Y%m%d_%H%M%S')}.csv"
        
        with open(filename, 'w') as f:
            # Define headers
            headers = [
                'asset_id', 'hostname', 'serial_number', 'manufacturer', 'model_id',
                'flag_timestamp', 'flag_tech', 'flag_notes', 'status', 'site'
            ]
            
            # Write headers
            f.write(','.join([f'"{h}"' for h in headers]) + '\n')
            
            # Write data
            for item in inventory:
                # Format data
                row_data = {
                    'asset_id': item.get('asset_id', ''),
                    'hostname': item.get('hostname', ''),
                    'serial_number': item.get('serial_number', ''),
                    'manufacturer': item.get('manufacturer', ''),
                    'model_id': item.get('model_id', ''),
                    'flag_timestamp': item.get('flag_timestamp', ''),
                    'flag_tech': item.get('flag_tech', ''),
                    'flag_notes': item.get('flag_notes', ''),
                    'status': item.get('status', ''),
                    'site': item.get('site', '')
                }
                
                # Write row
                row = [f'"{str(row_data.get(h, ""))}"' for h in headers]
                f.write(','.join(row) + '\n')
        
        # Report success
        asset_count = len(inventory)
        messagebox.showinfo("Export Complete", 
                           f"Flagged assets exported to {filename}\n\n"
                           f"Total flagged assets: {asset_count}")