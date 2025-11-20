import tkinter as tk
from tkinter import ttk, messagebox
import logging
import os

from ui.scan_tab import ScanTab
from ui.inventory_tab import InventoryTab 
from ui.all_benches_tab import AllBenchesTab
from ui.checked_out_tab import CheckedOutTab
from ui.history_tab import HistoryTab
from ui.flagged_assets_tab import FlaggedAssetsTab
from ui.daas_expiring_tab import DaasExpiringTab
from models.database import InventoryDatabase

# Setup logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s',
                   filename='inventory_app.log', filemode='a')
logger = logging.getLogger(__name__)

# Add console handler
console = logging.StreamHandler()
console.setLevel(logging.INFO)
logger.addHandler(console)

class InventoryApp:
    def __init__(self, root, config=None):
        self.root = root
        self.root.title("IT Bench Inventory Manager")
        # self.root.geometry("1000x600") # Keep this commented/removed

        # Store the configuration
        self.config = config or {}

        # Get the current directory (Note: InventoryDatabase likely handles its own path)
        current_dir = os.path.dirname(os.path.abspath(__file__))
        db_path = os.path.join(current_dir, "inventory.db") # This path might not be directly used by InventoryDatabase

        logger.info(f"Initializing database connection using db_config.json")
        self.db = InventoryDatabase('db_config.json') # Assumes db_config.json is in the right place

        # ---> Configure Root Window Grid <---
        self.root.rowconfigure(0, weight=1)  # Row 0 for the notebook (expands)
        self.root.rowconfigure(1, weight=0)  # Row 1 for the status bar (fixed height)
        self.root.columnconfigure(0, weight=1) # Column 0 takes all width

        # Create notebook (tabbed interface)
        self.notebook = ttk.Notebook(root)
        # self.notebook.pack(fill='both', expand=True) # REMOVED pack
        self.notebook.grid(row=0, column=0, sticky="nsew") # Place notebook in grid

        # ---> Define Global Styles Here <---
        style = ttk.Style()
        # Green Button
        style.configure("Green.TButton", background="#4CAF50", foreground="black", bordercolor="#4CAF50")
        style.map("Green.TButton",
                background=[("active", "#45a049")],
                bordercolor=[("active", "#45a049")])
        # Red Button
        style.configure("Red.TButton", background="#F44336", foreground="black", bordercolor="#F44336")
        style.map("Red.TButton",
                background=[("active", "#d32f2f")],
                bordercolor=[("active", "#d32f2f")])
        # Orange Button
        style.configure("Orange.TButton", background="#FF9800", foreground="black", bordercolor="#FF9800")
        style.map("Orange.TButton",
                background=[("active", "#e68a00")],
                bordercolor=[("active", "#e68a00")])
        # Purple Button for DaaS
        style.configure("Purple.TButton", background="#9C27B0", foreground="black", bordercolor="#9C27B0")
        style.map("Purple.TButton",
                background=[("active", "#7B1FA2")],
                bordercolor=[("active", "#7B1FA2")])
        # ---> End of Global Styles <---

        # Initialize all tabs
        self.scan_tab = ScanTab(self.notebook, self.db, self, self.config)
        self.inventory_tab = InventoryTab(self.notebook, self.db, self, self.config)
        self.all_benches_tab = AllBenchesTab(self.notebook, self.db, self, self.config)
        self.checked_out_tab = CheckedOutTab(self.notebook, self.db, self, self.config)
        self.flagged_assets_tab = FlaggedAssetsTab(self.notebook, self.db, self, self.config)
        self.daas_expiring_tab = DaasExpiringTab(self.notebook, self.db, self, self.config)
        self.history_tab = HistoryTab(self.notebook, self.db, self, self.config)

        # Add tabs to notebook in the right order (excluding flagged if it failed)
        self.notebook.add(self.scan_tab.frame, text="Scan")
        self.notebook.add(self.inventory_tab.frame, text="Your Bench")
        self.notebook.add(self.all_benches_tab.frame, text="All Benches")
        self.notebook.add(self.checked_out_tab.frame, text="Out")
        self.notebook.add(self.flagged_assets_tab.frame, text="Flagged")
        self.notebook.add(self.daas_expiring_tab.frame, text="DaaS Expiring")
        self.notebook.add(self.history_tab.frame, text="Recent History")

        # Add status bar
        self.create_status_bar() # This method defines self.status_bar
        # self.status_bar.pack(side=tk.BOTTOM, fill=tk.X) # REMOVED pack
        self.status_bar.grid(row=1, column=0, sticky="ew") # Place status_bar in grid

        # Bind tab change event to refresh the selected tab
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_changed)

        # Schedule regular updates of the status indicators
        self.update_status_indicator()

        # Update lease expiry flags on startup
        self.db.update_all_expiry_flags()

        logger.info(f"Application initialized for site: {self.config.get('site', 'Unknown')}")
    
    def create_status_bar(self):
            """Create a status bar frame (widgets inside still use pack relative to this frame)"""
            # self.status_bar frame is placed into root using grid in __init__
            self.status_bar = ttk.Frame(self.root, relief=tk.SUNKEN, padding=(2, 2))

            # Widgets INSIDE the status bar still use pack relative to self.status_bar
            style = ttk.Style()
            style.configure("Green.TLabel", background="#4CAF50", foreground="white")
            style.configure("Orange.TLabel", background="#FF9800", foreground="white")

            indicator_frame = ttk.Frame(self.status_bar)
            indicator_frame.pack(side=tk.LEFT, padx=5) # Pack INSIDE status_bar

            # Use a neutral initial message
            self.db_status_text = tk.StringVar(value="Checking status...")
            # --- CORRECTED LABEL ARGUMENTS ---
            self.db_status_label = ttk.Label(indicator_frame, textvariable=self.db_status_text, padding=(5, 2), width=30)
            # --- END CORRECTION ---
            self.db_status_label.pack(side=tk.LEFT) # Pack INSIDE indicator_frame
            # Sync button setup
            self.sync_button = ttk.Button(indicator_frame, text="Sync to DB", command=self.sync_to_db)
            self.sync_button.pack(side=tk.LEFT, padx=5) # Pack INSIDE indicator_frame
            # Initially hide sync button if not using local DB (check happens in update_status_indicator)
            if not hasattr(self.db, 'using_local') or not self.db.using_local:
                self.sync_button.pack_forget()

            # Secondary status label setup
            self.secondary_status = tk.StringVar(value="")
            self.secondary_status_label = ttk.Label(self.status_bar, textvariable=self.secondary_status, padding=(5, 2))
            self.secondary_status_label.pack(side=tk.RIGHT, padx=5) # Pack INSIDE status_bar
    
    def update_status_indicator(self):
        """Update the database connection status indicator"""
        # Check current status
        if self.db.using_local:
            self.db_status_text.set("Using Local Database Backup")
            self.db_status_label.configure(style="Orange.TLabel")
            
            # Show sync button
            if not self.sync_button.winfo_ismapped():
                self.sync_button.pack(side=tk.LEFT, padx=5)
                
            # Show if there are pending changes
            if self.db.pending_sync:
                self.secondary_status.set("Pending changes to sync")
            else:
                self.secondary_status.set("Local database is up to date")
        else:
            self.db_status_text.set("Connected to PostgreSQL")
            self.db_status_label.configure(style="Green.TLabel")
            
            # Hide sync button
            if self.sync_button.winfo_ismapped():
                self.sync_button.pack_forget()
                
            self.secondary_status.set("")
        
        # Schedule next update
        self.root.after(5000, self.update_status_indicator)  # Update every 5 seconds
    
    def sync_to_db(self):
        """Attempt to sync with central database"""
        self.secondary_status.set("Syncing to central database...")
        self.root.update_idletasks()

        was_local = self.db.using_local # Remember if we were previously offline

        try:
            if was_local:
                logger.info("Attempting to re-initialize PostgreSQL connection pools...")
                try:
                    self.db._initialize_pg_connection_pools()
                    logger.info("PostgreSQL pools re-initialized successfully.")
                    self.db.using_local = False # Now we can tentatively switch
                except Exception as pool_init_error:
                    logger.error(f"Failed to re-initialize PostgreSQL pools: {pool_init_error}")
                    self.db.using_local = True # Ensure we stay local
                    raise pool_init_error # Re-raise the error to be caught below
            if not self.db.using_local:
                conn = self.db._get_pg_connection(write=True)
                self.db.release_connection(conn) # Test connection is successful

                self.db._sync_to_server() # _sync_to_server likely checks self.pending_sync
                if self.db.pending_sync:
                    self.secondary_status.set("Partial sync completed. Some changes still pending.")
                else:
                    self.secondary_status.set("Sync completed successfully!")
                    if was_local:
                        self.db_status_text.set("Connected to PostgreSQL")
                        self.db_status_label.configure(style="Green.TLabel")
                        if self.sync_button.winfo_ismapped():
                            self.sync_button.pack_forget()
            else:
                 raise Exception("Still unable to connect after potential re-initialization.")

        except Exception as e:
            if not self.db.using_local and was_local: # Only reset if we attempted to switch
                 self.db.using_local = True

            error_msg = str(e)
            logger.error(f"Sync failed: {error_msg}")
            self.secondary_status.set(f"Sync failed: {error_msg[:50]}...")

            messagebox.showerror("Sync Failed",
                               f"Could not connect or sync with central database.\n\nError: {error_msg}\n\nChanges remain in local database.")
        finally:
             self.update_status_indicator()
    
    def on_tab_changed(self, event):
        """Handle tab change event by refreshing the selected tab"""
        selected_tab = self.notebook.select()
        tab_id = self.notebook.index(selected_tab)
        
        logger.info(f"Tab changed to index: {tab_id}")
        
        # Refresh the appropriate tab based on its index
        if tab_id == 1:  # Your Bench tab
            logger.info("Refreshing Your Bench tab")
            try:
                self.inventory_tab.refresh_inventory()
            except Exception as e:
                logger.error(f"Error refreshing Your Bench tab: {e}")
        elif tab_id == 2:  # All Benches tab
            logger.info("Refreshing All Benches tab")
            try:
                self.all_benches_tab.refresh_inventory()
            except Exception as e:
                logger.error(f"Error refreshing All Benches tab: {e}")
        elif tab_id == 3:  # Checked Out tab
            logger.info("Refreshing checked out tab")
            try:
                self.checked_out_tab.refresh_inventory()
            except Exception as e:
                logger.error(f"Error refreshing checked out tab: {e}")
        elif tab_id == 4:  # Flagged Assets tab
            logger.info("Refreshing flagged assets tab")
            try:
                self.flagged_assets_tab.refresh_inventory()
            except Exception as e:
                logger.error(f"Error refreshing flagged assets tab: {e}")
        elif tab_id == 5:  # DaaS Expiring tab
            logger.info("Refreshing DaaS expiring tab")
            try:
                self.daas_expiring_tab.refresh_inventory()
            except Exception as e:
                logger.error(f"Error refreshing DaaS expiring tab: {e}")
        elif tab_id == 6:  # Recent History tab - update index
            logger.info("Refreshing history tab")
            try:
                self.history_tab.refresh_history()
            except Exception as e:
                logger.error(f"Error refreshing history tab: {e}")
        
    def refresh_all(self):
        """Refresh all tabs"""
        logger.info("Refreshing all tabs")
        try:
            self.inventory_tab.refresh_inventory()
            logger.info("Your Bench tab refreshed")
        except Exception as e:
            logger.error(f"Error refreshing Your Bench tab: {e}")
        
        try:
            self.all_benches_tab.refresh_inventory()
            logger.info("All Benches tab refreshed")
        except Exception as e:
            logger.error(f"Error refreshing All Benches tab: {e}")
        
        try:
            self.checked_out_tab.refresh_inventory()
            logger.info("Checked out tab refreshed")
        except Exception as e:
            logger.error(f"Error refreshing checked out tab: {e}")
        
        try:
            self.flagged_assets_tab.refresh_inventory()
            logger.info("Flagged assets tab refreshed")
        except Exception as e:
            logger.error(f"Error refreshing flagged assets tab: {e}")
        
        try:
            self.daas_expiring_tab.refresh_inventory()
            logger.info("DaaS expiring tab refreshed")
        except Exception as e:
            logger.error(f"Error refreshing DaaS expiring tab: {e}")
        
        try:
            self.history_tab.refresh_history()
            logger.info("History tab refreshed")
        except Exception as e:
            logger.error(f"Error refreshing history tab: {e}")

    def show_local_save_notification(self):
        """Show a notification that changes were saved locally"""
        if self.db.using_local and self.db.pending_sync:
            # Get pending changes count
            count = self.db.get_pending_changes_count()
            if count > 0:
                messagebox.showinfo("Local Changes Saved", 
                                f"{count} change(s) have been saved to the local database.\n\n"
                                "These changes will be automatically synced when the central database becomes available.\n\n"
                                "You can manually sync by clicking the 'Sync to DB' button in the status bar.")