import tkinter as tk
from tkinter import ttk, messagebox
import logging

from ui.utils import add_context_menu
from ui.dialogs import show_asset_details, create_properly_sized_dialog
from ui.dialogs import format_timestamp

logger = logging.getLogger(__name__)

class HistoryTab:
    def __init__(self, parent, db, app, config=None):
        self.db = db
        self.app = app
        self.config = config or {}
        self.frame = ttk.Frame(parent)
        self.setup_ui()
        
    def setup_ui(self):
        """Set up the history tab UI"""
        # Frame for history
        history_frame = ttk.Frame(self.frame)
        history_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        # Controls
        control_frame = ttk.Frame(history_frame)
        control_frame.pack(fill="x", padx=5, pady=5)
        
        ttk.Button(control_frame, text="Refresh", 
                  command=self.refresh_history).pack(side="left", padx=5)
        ttk.Button(control_frame, text="Search Older Records", 
                  command=self.search_full_history).pack(side="left", padx=5)
        
        # Add toggle for showing records from all sites
        self.show_all_sites = tk.BooleanVar(value=False)
        ttk.Checkbutton(control_frame, text="Show All Sites", 
                       variable=self.show_all_sites, 
                       command=self.refresh_history).pack(side="left", padx=20)
        
        # Site indicator
        site = self.config.get('site', 'Unknown')
        site_label = ttk.Label(control_frame, text=f"Site: {site}", font=("", 10, "bold"))
        site_label.pack(side="right", padx=20)
        
        # Treeview for history list
        columns = ("timestamp", "asset_id", "serial", "status", "tech", "notes", "site")
        self.history_tree = ttk.Treeview(history_frame, columns=columns, show="headings")
        
        # Define headings with sorting
        self.history_tree.heading("timestamp", text="Date/Time", command=lambda: self.treeview_sort_column(self.history_tree, "timestamp", False))
        self.history_tree.heading("asset_id", text="Asset Tag", command=lambda: self.treeview_sort_column(self.history_tree, "asset_id", False))
        self.history_tree.heading("serial", text="Serial Number", command=lambda: self.treeview_sort_column(self.history_tree, "serial", False))
        self.history_tree.heading("status", text="Action", command=lambda: self.treeview_sort_column(self.history_tree, "status", False))
        self.history_tree.heading("tech", text="Technician", command=lambda: self.treeview_sort_column(self.history_tree, "tech", False))
        self.history_tree.heading("notes", text="Notes", command=lambda: self.treeview_sort_column(self.history_tree, "notes", False))
        self.history_tree.heading("site", text="Site", command=lambda: self.treeview_sort_column(self.history_tree, "site", False))
        
        # Define columns
        self.history_tree.column("timestamp", width=150)
        self.history_tree.column("asset_id", width=100)
        self.history_tree.column("serial", width=150)
        self.history_tree.column("status", width=80)
        self.history_tree.column("tech", width=100)
        self.history_tree.column("notes", width=200)
        self.history_tree.column("site", width=80)
        
        # Add scrollbar
        scrollbar = ttk.Scrollbar(history_frame, orient="vertical", command=self.history_tree.yview)
        self.history_tree.configure(yscrollcommand=scrollbar.set)
        
        self.history_tree.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Bind double-click to view details
        self.history_tree.bind("<Double-1>", self.on_history_double_click)
        
        # Add context menu
        add_context_menu(self.history_tree)
        
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
    
    def on_history_double_click(self, event):
        """Handle double-click on history item"""
        item = self.history_tree.identify('item', event.x, event.y)
        if item:
            asset_id = self.history_tree.item(item, "values")[1]
            show_asset_details(self.db, asset_id, self.config)  # Pass config here

    def refresh_history(self):
        """Refresh the recent history view"""
        # Clear existing items
        for item in self.history_tree.get_children():
            self.history_tree.delete(item)
        
        # Get recent history
        history = self.db.get_recent_history()
        
        # Filter by site if needed
        show_all_sites = self.show_all_sites.get() if hasattr(self, 'show_all_sites') else False
        current_site = self.config.get('site')
        
        if current_site and not show_all_sites:
            # Only keep items from the current site or with NULL site
            filtered_history = []
            for item in history:
                item_site = item.get('site')
                # Include items that match current site OR have no site (for backward compatibility)
                if item_site == current_site or item_site is None or item_site == '' or item_site == 'Unknown':
                    filtered_history.append(item)
            history = filtered_history
        
        # Populate treeview
        for item in history:
            # Add tags for site highlighting
            tags = ()
            if item.get('site') == current_site:
                tags = ('current_site',)
            
            self.history_tree.insert("", "end", values=(
                format_timestamp(item.get('timestamp', '')),
                item.get('asset_id', ''),
                item.get('serial_number', ''),
                item.get('status', ''),
                item.get('tech_name', ''),
                item.get('notes', ''),
                item.get('site', 'Unknown')
            ), tags=tags)
        
        # Configure tag for current site items (bold)
        self.history_tree.tag_configure('current_site', font=('', 9, 'bold'))
    
    def search_full_history(self):
        """Open dialog to search complete history"""
        search_window = create_properly_sized_dialog("Search Full History", 500, 200)
        
        # Add content frame
        content_frame = ttk.Frame(search_window)
        content_frame.pack(fill="both", expand=True, padx=10, pady=10)
        
        ttk.Label(content_frame, text="Enter Asset Tag or Serial Number:").pack(pady=10)
        
        search_entry = ttk.Entry(content_frame, width=30)
        search_entry.pack(pady=5)
        search_entry.focus()
        
        # Site filter
        site_frame = ttk.Frame(content_frame)
        site_frame.pack(fill="x", pady=10)
        
        show_all_sites = tk.BooleanVar(value=False)
        site_check = ttk.Checkbutton(site_frame, text="Show results from all sites", 
                                    variable=show_all_sites)
        site_check.pack(side="left")
        
        current_site = self.config.get('site', 'Unknown')
        ttk.Label(site_frame, text=f"Current site: {current_site}").pack(side="left", padx=10)
        
        # Add context menu
        add_context_menu(search_entry)
        
        def perform_search():
            search_term = search_entry.get().strip()
            if not search_term:
                return
                
            # Search database - For asset tags, try uppercase version as well
            results = self.db.search_asset_history(search_term)
            
            # If no results found and looks like an asset tag (starts with GF-), try uppercase
            if not results and search_term.lower().startswith("gf-"):
                uppercase_term = search_term.upper()
                if uppercase_term != search_term:  # Only search again if it's different
                    results = self.db.search_asset_history(uppercase_term)
            
            # If still no results, try lowercase for serial numbers
            if not results:
                lowercase_term = search_term.lower()
                if lowercase_term != search_term:  # Only search again if it's different
                    results = self.db.search_asset_history(lowercase_term)
            
            # Filter by site if needed
            if not show_all_sites.get() and current_site:
                results = [item for item in results if item.get('site') == current_site or item.get('site') is None]
            
            if not results:
                messagebox.showinfo("No Results", "No history found for this search term")
                return
            
            # Rest of the function remains the same...
            # Create results window
            results_window = create_properly_sized_dialog(f"History for {search_term}", 800, 400)
            
            # Create treeview for results
            columns = ("timestamp", "asset_id", "status", "tech", "notes", "site")
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
            results_tree.heading("site", text="Site", 
                                command=lambda: self.treeview_sort_column(results_tree, "site", False))
            
            results_tree.column("timestamp", width=150)
            results_tree.column("asset_id", width=100)
            results_tree.column("status", width=80)
            results_tree.column("tech", width=100)
            results_tree.column("notes", width=250)
            results_tree.column("site", width=80)
            
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
                    show_asset_details(self.db, asset_id)
            
            results_tree.bind("<Double-1>", on_double_click)
            
            # Populate results
            for item in results:
                # Add tags for site highlighting
                tags = ()
                if item.get('site') == current_site:
                    tags = ('current_site',)
                
                results_tree.insert("", "end", values=(
                    format_timestamp(item.get('timestamp', '')),
                    item.get('asset_id', ''),
                    item.get('status', ''),
                    item.get('tech_name', ''),
                    item.get('notes', ''),
                    item.get('site', 'Unknown')
                ), tags=tags)
            
            # Configure tag for current site items (bold)
            results_tree.tag_configure('current_site', font=('', 9, 'bold'))
        
        ttk.Button(content_frame, text="Search", command=perform_search).pack(pady=10)
        
        # Bind Enter key to search
        search_entry.bind("<Return>", lambda e: perform_search())