import psycopg2
from psycopg2 import pool
from psycopg2.extras import DictCursor
import os
import sqlite3
from datetime import datetime
import logging
import time
import random
import json
import threading
import sys

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class InventoryDatabase:
    def __init__(self, config_file=None):
            """Initialize the database with primary PostgreSQL and local SQLite backup"""
            # Default configuration
            self.db_config = {
                'primary': {
                    'dbname': 'inventory',
                    'user': 'inventory_user',
                    'password': 'secure_password',  # In production, use secure password management
                    'host': 'primary-db-server.example.com',
                    'port': '5432'
                },
                'replicas': [
                    {
                        'dbname': 'inventory',
                        'user': 'inventory_user',
                        'password': 'secure_password',
                        'host': 'replica1-db-server.example.com',
                        'port': '5432'
                    },
                    # Add more replicas as needed
                ]
            }

            # Initialize status flags
            self.using_local = False
            self.pending_sync = False
            self.last_sync_time = None
            self.sqlite_conn = None

            # Set up local SQLite database path
            if getattr(sys, 'frozen', False):
                # Running as frozen executable (PyInstaller package)
                base_dir = os.path.dirname(sys.executable)
            else:
                # Running as script
                base_dir = os.path.dirname(os.path.abspath(__file__))
                
            self.local_db_path = os.path.join(base_dir, "local_inventory.db")
            logger.info(f"Local backup database path: {self.local_db_path}")

            # Load configuration from file if provided
            if config_file and os.path.exists(config_file):
                self._load_config(config_file)

            # Initialize the connection pools as None first
            self.primary_pool = None
            self.replica_pools = []

            # Try to initialize PostgreSQL connection pools
            try:
                logger.info("Initializing PostgreSQL connection pools")
                self._initialize_pg_connection_pools()
                self.using_local = False

                # --- NEW: Sync central to local on successful connection ---
                logger.info("Attempting initial sync from central DB to local DB...")
                try:
                    # Ensure SQLite is initialized before syncing TO it
                    logger.info("Ensuring local SQLite backup database schema is initialized...")
                    self._initialize_sqlite_database()
                    # Call the sync function (assuming it exists)
                    self._sync_central_to_local()
                    logger.info("Initial sync from central to local completed successfully.")
                except Exception as sync_error:
                    logger.warning(f"Initial sync from central to local failed: {sync_error}")
                    logger.warning("Proceeding without initial sync. Local DB might be stale.")
                    # Decide if you want to force local mode if initial sync fails, or just log it.
                    # Forcing local mode:
                    # self.using_local = True
                    # self.pending_sync = False # No server changes to sync *back* yet
                    # logger.warning("Forcing local mode due to failed initial sync.")
                # --- END NEW SECTION ---

            except Exception as e:
                logger.warning(f"Could not connect to PostgreSQL server: {e}")
                logger.warning("Will use local SQLite database...")
                self.using_local = True
                self.pending_sync = False
                # Ensure SQLite is initialized if we fell back here
                logger.info("Initializing local SQLite backup database (fallback)...")
                self._initialize_sqlite_database() # Creates schema if needed


            # Initialize schema/tables in the *currently active* database
            logger.info(f"Initializing database schema for {'PostgreSQL' if not self.using_local else 'SQLite'}...")
            self._initialize_database()

            # Start background connectivity checker
            self._start_connectivity_checker()
  
    def _load_config(self, config_file):
        """Load database configuration from file"""
        import json
        try:
            with open(config_file, 'r') as f:
                self.db_config = json.load(f)
            logger.info(f"Loaded database configuration from {config_file}")
        except Exception as e:
            logger.error(f"Error loading config file: {str(e)}")
    
    def _initialize_pg_connection_pools(self):
        """Initialize connection pools for primary and replica databases"""
        # Create primary connection pool
        primary_dsn = " ".join([f"{k}={v}" for k, v in self.db_config['primary'].items()])
        self.primary_pool = pool.ThreadedConnectionPool(
            minconn=3,    # More minimum connections
            maxconn=20,   # Higher maximum for remote users
            dsn=primary_dsn
        )
        
        # Create replica connection pools
        self.replica_pools = []
        for replica_config in self.db_config.get('replicas', []):
            replica_dsn = " ".join([f"{k}={v}" for k, v in replica_config.items()])
            replica_pool = pool.ThreadedConnectionPool(
                minconn=1,
                maxconn=5,
                dsn=replica_dsn
            )
            self.replica_pools.append(replica_pool)
    
    def _initialize_sqlite_database(self):
        """Initialize the SQLite backup database"""
        conn = None
        try:
            conn = sqlite3.connect(self.local_db_path)
            cursor = conn.cursor()
            
            # Create assets table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS assets (
                    asset_id TEXT PRIMARY KEY,
                    serial_number TEXT,
                    hostname TEXT,
                    operational_status TEXT,
                    install_status TEXT,
                    location TEXT,
                    ci_region TEXT,
                    owned_by TEXT,
                    assigned_to TEXT,
                    comments TEXT,
                    manufacturer TEXT,
                    model_id TEXT,
                    model_description TEXT,
                    vendor TEXT,
                    warranty_expiration TEXT,
                    os TEXT,
                    os_version TEXT,
                    cmdb_url TEXT,
                    last_updated TIMESTAMP,
                    flag_status INTEGER DEFAULT 0,
                    flag_notes TEXT,
                    flag_timestamp TIMESTAMP,
                    flag_tech TEXT,
                    lease_start_date TEXT,
                    lease_maturity_date TEXT,
                    expiry_flag_status INTEGER DEFAULT 0
                )
            ''')
            
            # Check if flag_status column exists
            try:
                cursor.execute("SELECT flag_status FROM assets LIMIT 1")
            except sqlite3.OperationalError:
                # Column doesn't exist, add it
                cursor.execute("ALTER TABLE assets ADD COLUMN flag_status INTEGER DEFAULT 0")
                cursor.execute("ALTER TABLE assets ADD COLUMN flag_notes TEXT")
                cursor.execute("ALTER TABLE assets ADD COLUMN flag_timestamp TIMESTAMP")
                cursor.execute("ALTER TABLE assets ADD COLUMN flag_tech TEXT")
            
            # Check if lease columns exist
            try:
                cursor.execute("SELECT lease_start_date FROM assets LIMIT 1")
            except sqlite3.OperationalError:
                # Lease columns don't exist, add them
                cursor.execute("ALTER TABLE assets ADD COLUMN lease_start_date TEXT")
                cursor.execute("ALTER TABLE assets ADD COLUMN lease_maturity_date TEXT")
                cursor.execute("ALTER TABLE assets ADD COLUMN expiry_flag_status INTEGER DEFAULT 0")
            
            # Create scan_history table with site field
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS scan_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    asset_id TEXT,
                    status TEXT,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    notes TEXT,
                    tech_name TEXT,
                    site TEXT,
                    FOREIGN KEY (asset_id) REFERENCES assets(asset_id)
                )
            ''')
            
            # Create related_items table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS related_items (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    parent_asset_id TEXT,
                    serial_number TEXT UNIQUE,
                    item_type TEXT,
                    notes TEXT,
                    FOREIGN KEY (parent_asset_id) REFERENCES assets(asset_id)
                )
            ''')
            
            # Create sync_queue table for tracking changes to sync
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS sync_queue (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    operation TEXT,
                    table_name TEXT,
                    data TEXT,
                    timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP
                )
            ''')
            
            conn.commit()
            logger.info("SQLite database tables initialized")
        except Exception as e:
            logger.error(f"Error initializing SQLite database: {str(e)}")
        finally:
            if conn:
                conn.close()
    
    def _initialize_database(self):
        """Create the database tables if they don't exist"""
        conn = None
        try:
            conn = self.get_connection(write=True)
            cursor = conn.cursor()
            
            # Create tables if they don't exist
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS assets (
                    asset_id TEXT PRIMARY KEY,
                    serial_number TEXT,
                    hostname TEXT,
                    operational_status TEXT,
                    install_status TEXT,
                    location TEXT,
                    ci_region TEXT,
                    owned_by TEXT,
                    assigned_to TEXT,
                    comments TEXT,
                    manufacturer TEXT,
                    model_id TEXT,
                    model_description TEXT,
                    vendor TEXT,
                    warranty_expiration TEXT,
                    os TEXT,
                    os_version TEXT,
                    cmdb_url TEXT,
                    last_updated TIMESTAMP WITH TIME ZONE
                )
            ''')
            
            # Check if flag_status column exists in PostgreSQL
            if not self.using_local:
                # PostgreSQL - check if column exists
                cursor.execute('''
                    SELECT column_name 
                    FROM information_schema.columns 
                    WHERE table_name='assets' AND column_name='flag_status'
                ''')
                flag_column_exists = cursor.fetchone() is not None
                
                # Add flag columns if they don't exist
                if not flag_column_exists:
                    try:
                        cursor.execute("ALTER TABLE assets ADD COLUMN flag_status BOOLEAN DEFAULT FALSE")
                        cursor.execute("ALTER TABLE assets ADD COLUMN flag_notes TEXT")
                        cursor.execute("ALTER TABLE assets ADD COLUMN flag_timestamp TIMESTAMP WITH TIME ZONE")
                        cursor.execute("ALTER TABLE assets ADD COLUMN flag_tech TEXT")
                        conn.commit()
                        logger.info("Added flag columns to assets table")
                    except Exception as e:
                        logger.error(f"Error adding flag columns: {str(e)}")
                        conn.rollback()
                
                # Check if lease columns exist
                cursor.execute('''
                    SELECT column_name 
                    FROM information_schema.columns 
                    WHERE table_name='assets' AND column_name='lease_start_date'
                ''')
                lease_column_exists = cursor.fetchone() is not None
                
                # Add lease columns if they don't exist
                if not lease_column_exists:
                    try:
                        cursor.execute("ALTER TABLE assets ADD COLUMN lease_start_date TEXT")
                        cursor.execute("ALTER TABLE assets ADD COLUMN lease_maturity_date TEXT")
                        cursor.execute("ALTER TABLE assets ADD COLUMN expiry_flag_status BOOLEAN DEFAULT FALSE")
                        conn.commit()
                        logger.info("Added lease columns to assets table")
                    except Exception as e:
                        logger.error(f"Error adding lease columns: {str(e)}")
                        conn.rollback()
            
            # Check if site column exists in scan_history
            if not self.using_local:
                # PostgreSQL - check if column exists
                cursor.execute('''
                    SELECT column_name 
                    FROM information_schema.columns 
                    WHERE table_name='scan_history' AND column_name='site'
                ''')
                site_column_exists = cursor.fetchone() is not None
                
                # Create scan_history table if it doesn't exist, or add site column if needed
                if not site_column_exists:
                    try:
                        # First check if the table exists
                        cursor.execute('''
                            SELECT to_regclass('public.scan_history')
                        ''')
                        table_exists = cursor.fetchone()[0] is not None
                        
                        if not table_exists:
                            logger.info("Creating scan_history table with site column")
                            cursor.execute('''
                                CREATE TABLE IF NOT EXISTS scan_history (
                                    id SERIAL PRIMARY KEY,
                                    asset_id TEXT,
                                    status TEXT,
                                    timestamp TIMESTAMP WITH TIME ZONE DEFAULT CURRENT_TIMESTAMP,
                                    notes TEXT,
                                    tech_name TEXT,
                                    site TEXT,
                                    FOREIGN KEY (asset_id) REFERENCES assets(asset_id)
                                )
                            ''')
                        else:
                            # Table exists but missing site column, add it
                            logger.info("Adding site column to scan_history table")
                            cursor.execute('''
                                ALTER TABLE scan_history ADD COLUMN site TEXT
                            ''')
                        
                        conn.commit()
                    except Exception as e:
                        logger.error(f"Error updating scan_history table: {str(e)}")
                        conn.rollback()
                        raise
            else:
                # SQLite - table creation is already handled in _initialize_sqlite_database
                pass
            
            # Create related_items table
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS related_items (
                    id SERIAL PRIMARY KEY,
                    parent_asset_id TEXT,
                    serial_number TEXT UNIQUE,
                    item_type TEXT,
                    notes TEXT,
                    FOREIGN KEY (parent_asset_id) REFERENCES assets(asset_id)
                )
            ''')
            
            conn.commit()
            logger.info("Database tables initialized")
        except Exception as e:
            logger.error(f"Error initializing database: {str(e)}")
            if conn:
                conn.rollback()
        finally:
            if conn:
                self.release_connection(conn)
    
    def _get_pg_connection(self, write=False):
        """Get a PostgreSQL connection from the pool with automatic failover"""
        if write or not self.replica_pools:
            # For write operations or if no replicas available, use primary
            if not self.primary_pool:
                raise Exception("Primary database pool not initialized")
            return self.primary_pool.getconn()
        else:
            # For read operations, randomly select a replica for load balancing
            # Start with replica if available
            try:
                replica_pool = random.choice(self.replica_pools)
                return replica_pool.getconn()
            except Exception:
                # Failover to primary
                return self.primary_pool.getconn()
    
    def _get_sqlite_connection(self):
        """Get a shared connection to the SQLite backup database."""
        try:
            # Check if connection exists and is usable
            if self.sqlite_conn is None:
                logger.info("Creating new shared SQLite connection.")
                self.sqlite_conn = sqlite3.connect(self.local_db_path, check_same_thread=False) # Allow sharing across threads if needed by background checker later
                # Optionally set row factory or other settings here if needed
                # self.sqlite_conn.row_factory = sqlite3.Row
            # Minimal check if connection seems alive (can execute a simple query)
            self.sqlite_conn.execute("SELECT 1")
            return self.sqlite_conn
        except (sqlite3.ProgrammingError, sqlite3.OperationalError) as e:
            # Connection might be closed or broken, try reconnecting
            logger.warning(f"Shared SQLite connection issue ({e}), attempting to reconnect.")
            try:
                if self.sqlite_conn:
                    self.sqlite_conn.close()
            except Exception:
                pass # Ignore errors closing a potentially broken connection
            self.sqlite_conn = sqlite3.connect(self.local_db_path, check_same_thread=False)
            return self.sqlite_conn
        except Exception as e:
            logger.error(f"Failed to get SQLite connection: {e}")
            raise # Re-raise critical errors
    
    def get_connection(self, write=False):
        """Get database connection with failover to local SQLite"""
        # Try PostgreSQL first
        if not self.using_local:
            try:
                return self._get_pg_connection(write)
            except Exception as e:
                logger.warning(f"PostgreSQL connection failed: {e}")
                # Switch to local mode
                self.using_local = True
                self.pending_sync = True
        
        # Use SQLite as fallback
        try:
            return self._get_sqlite_connection()
        except Exception as e:
            logger.error(f"All database connections failed: {e}")
            raise
    
    def release_connection(self, conn):
        """Return a connection to its pool (PostgreSQL) or do nothing (SQLite)."""
        if self.using_local:
            return
        try:
            for pool_obj in [self.primary_pool] + self.replica_pools:
                try:
                    if pool_obj:
                        pool_obj.putconn(conn)
                        return
                except Exception:
                    continue
            logger.warning("Could not return PG connection to any pool")
        except Exception as e:
            logger.error(f"Error releasing PG connection: {str(e)}")

    def update_asset(self, asset_data):
        """Update or insert asset in database"""
        if not asset_data:
            logger.warning("Cannot update: No asset data provided")
            return False
            
        # Make sure we have an asset_tag
        asset_tag = asset_data.get('asset_tag', asset_data.get('asset_id'))
        if not asset_tag:
            logger.warning("Cannot update: No asset_tag provided")
            return False
        
        conn = None
        try:
            conn = self.get_connection(write=True)
            cursor = conn.cursor()
            
            # Check if asset exists
            if self.using_local:
                cursor.execute("SELECT asset_id FROM assets WHERE asset_id = ?", (asset_tag,))
            else:
                cursor.execute("SELECT asset_id FROM assets WHERE asset_id = %s", (asset_tag,))
            
            exists = cursor.fetchone()
            
            # Set last updated timestamp
            asset_data['last_updated'] = datetime.now()
            
            if exists:
                # Update existing asset
                set_clause = ", ".join([f"{key} = {'?' if self.using_local else '%s'}" for key in asset_data.keys() if key != 'asset_tag'])
                values = [asset_data[key] for key in asset_data.keys() if key != 'asset_tag']
                query = f"UPDATE assets SET {set_clause} WHERE asset_id = {'?' if self.using_local else '%s'}"
                values.append(asset_tag)
                cursor.execute(query, values)
                logger.info(f"Updated asset {asset_tag}")
            else:
                # Insert new asset - Map asset_tag to asset_id for the database schema
                asset_data_for_db = asset_data.copy()
                asset_data_for_db['asset_id'] = asset_tag
                if 'asset_tag' in asset_data_for_db and 'asset_tag' != 'asset_id':
                    del asset_data_for_db['asset_tag']
                    
                placeholders = ", ".join(['?' if self.using_local else '%s'] * len(asset_data_for_db))
                columns = ", ".join(asset_data_for_db.keys())
                query = f"INSERT INTO assets ({columns}) VALUES ({placeholders})"
                cursor.execute(query, list(asset_data_for_db.values()))
                logger.info(f"Inserted new asset {asset_tag}")
            
            conn.commit()
            
            # Record operation for sync if using local DB
            if self.using_local:
                self._record_operation("UPDATE" if exists else "INSERT", "assets", asset_data_for_db if not exists else asset_data)
                
            return True
        except Exception as e:
            logger.error(f"Database error updating asset: {str(e)}")
            if conn:
                conn.rollback()
            return False
        finally:
            if conn:
                self.release_connection(conn)
    
    def record_scan(self, asset_id, status, tech_name, notes="", site=None):
        """Record a scan event in history"""
        conn = None
        try:
            conn = self.get_connection(write=True)
            cursor = conn.cursor()
            
            scan_data = {
                'asset_id': asset_id,
                'status': status,
                'tech_name': tech_name,
                'notes': notes,
                'site': site,
                'timestamp': datetime.now()
            }

            db_timestamp = scan_data['timestamp']

            if self.using_local:
                cursor.execute('''
                    INSERT INTO scan_history (asset_id, status, tech_name, notes, site, timestamp)
                    VALUES (?, ?, ?, ?, ?, ?)
                ''', (asset_id, status, tech_name, notes, site, db_timestamp))
            else:
                cursor.execute('''
                    INSERT INTO scan_history (asset_id, status, tech_name, notes, site, timestamp)
                    VALUES (%s, %s, %s, %s, %s, %s)
                ''', (asset_id, status, tech_name, notes, site, db_timestamp))
            
            conn.commit()
            logger.info(f"Recorded scan for {asset_id}: status={status}, site={site}")
            
            # Record operation for sync if using local DB
            if self.using_local:
                # --- FIX: Convert datetime to string for JSON serialization ---
                # Create a copy or modify the dict for the sync queue
                sync_data = scan_data.copy()
                sync_data['timestamp'] = db_timestamp.isoformat() # Convert to ISO string format
                # --- END FIX ---
                self._record_operation("INSERT", "scan_history", sync_data) # Pass the modified dict

            return True
        except Exception as e:
            # Log the specific JSON serialization error if it happens here, though it should be caught inside _record_operation now
            logger.error(f"Error recording scan: {str(e)}", exc_info=True) # Added exc_info=True for more detail
            if conn:
                conn.rollback()
            return False
        finally:
            if conn:
                self.release_connection(conn)
    
    def get_asset_current_status(self, asset_id):
        """
        Get the current status (in/out) of an asset by looking ONLY
        at the latest record with status 'in' or 'out'.
        """
        conn = None
        try:
            conn = self.get_connection(write=False)  # Read operation
            cursor = conn.cursor()

            # Query for the most recent 'in' or 'out' record for this asset
            if self.using_local:
                # SQLite version
                cursor.execute('''
                    SELECT status, timestamp, tech_name, notes, site
                    FROM scan_history
                    WHERE asset_id = ? AND status IN ('in', 'out')
                    ORDER BY timestamp DESC
                    LIMIT 1
                ''', (asset_id,))

                # Get column names for dictionary conversion
                column_names = [description[0] for description in cursor.description]
                row = cursor.fetchone()
                result = dict(zip(column_names, row)) if row else None

            else:
                # PostgreSQL version
                cursor = conn.cursor(cursor_factory=DictCursor)
                cursor.execute('''
                    SELECT status, timestamp, tech_name, notes, site
                    FROM scan_history
                    WHERE asset_id = %s AND status IN ('in', 'out')
                    ORDER BY timestamp DESC
                    LIMIT 1
                ''', (asset_id,))
                result = cursor.fetchone() # Fetches as a dictionary

            if result:
                # If an 'in' or 'out' record was found, return it
                status_info = dict(result) if not self.using_local else result # Ensure it's a dict
                logger.info(f"Current status for asset {asset_id} (latest in/out): {status_info['status']}, site: {status_info.get('site')}")
                return status_info
            else:
                # No 'in' or 'out' history found for this asset, default to 'out'
                logger.warning(f"No 'in' or 'out' history found for asset {asset_id}, defaulting to 'out'.")
                return {
                    'status': 'out',
                    'timestamp': 'Unknown',
                    'tech_name': 'Unknown',
                    'notes': 'No check-in/out history found',
                    'site': None
                }
        except Exception as e:
            logger.error(f"Error getting current status for {asset_id}: {str(e)}")
            return {
                'status': 'unknown',
                'timestamp': 'Error',
                'tech_name': 'Error',
                'notes': f'Error: {str(e)}',
                'site': None
            }
        finally:
            if conn:
                self.release_connection(conn)
    
    def get_current_inventory(self, include_deleted=False):
        """Get all assets currently checked in"""
        conn = None
        try:
            conn = self.get_connection(write=False)  # Read operation
            cursor = conn.cursor()
            
            if self.using_local:
                # Modified query to find the latest check-in/check-out status
                query = '''
                    SELECT a.*, h1.timestamp as check_in_date, h1.site
                    FROM assets a
                    JOIN (
                        SELECT asset_id, MAX(id) as max_id
                        FROM scan_history
                        WHERE status IN ('in', 'out')
                        GROUP BY asset_id
                    ) latest ON a.asset_id = latest.asset_id
                    JOIN scan_history h1 ON latest.max_id = h1.id
                    WHERE h1.status = 'in'
                '''
                
                # Add filter for deleted assets if requested
                if not include_deleted:
                    query += " AND (a.operational_status IS NULL OR a.operational_status != 'DELETED')"
                    
                query += " ORDER BY h1.timestamp DESC"
                
                cursor.execute(query)
                
                # Get column names from cursor description
                column_names = [description[0] for description in cursor.description]
                
                # Convert rows to dictionaries
                result = []
                for row in cursor.fetchall():
                    row_dict = dict(zip(column_names, row))
                    result.append(row_dict)
            else:
                cursor = conn.cursor(cursor_factory=DictCursor)
                
                # Modified query to find the latest check-in/check-out status
                query = '''
                    SELECT a.*, h1.timestamp as check_in_date, h1.site
                    FROM assets a
                    JOIN (
                        SELECT asset_id, MAX(id) as max_id
                        FROM scan_history
                        WHERE status IN ('in', 'out')
                        GROUP BY asset_id
                    ) latest ON a.asset_id = latest.asset_id
                    JOIN scan_history h1 ON latest.max_id = h1.id
                    WHERE h1.status = 'in'
                '''
                
                # Add filter for deleted assets if requested
                if not include_deleted:
                    query += " AND (a.operational_status IS NULL OR a.operational_status != 'DELETED')"
                    
                query += " ORDER BY h1.timestamp DESC"
                
                cursor.execute(query)
                
                result = [dict(row) for row in cursor.fetchall()]
                
            return result
        except Exception as e:
            logger.error(f"Error getting current inventory: {str(e)}")
            return []
        finally:
            if conn:
                self.release_connection(conn)
    
    def get_checked_out_inventory(self, include_deleted=False):
        """Get all assets currently checked out"""
        conn = None
        try:
            conn = self.get_connection(write=False)  # Read operation
            cursor = conn.cursor()
            
            if self.using_local:
                # Modified query to find the latest check-in/check-out status
                query = '''
                    SELECT a.*, h1.timestamp as check_out_date, h1.site
                    FROM assets a
                    JOIN (
                        SELECT asset_id, MAX(id) as max_id
                        FROM scan_history
                        WHERE status IN ('in', 'out')
                        GROUP BY asset_id
                    ) latest ON a.asset_id = latest.asset_id
                    JOIN scan_history h1 ON latest.max_id = h1.id
                    WHERE h1.status = 'out'
                '''
                
                # Add filter for deleted assets if requested
                if not include_deleted:
                    query += " AND (a.operational_status IS NULL OR a.operational_status != 'DELETED')"
                    
                query += " ORDER BY h1.timestamp DESC"
                
                cursor.execute(query)
                
                # Get column names from cursor description
                column_names = [description[0] for description in cursor.description]
                
                # Convert rows to dictionaries
                result = []
                for row in cursor.fetchall():
                    row_dict = dict(zip(column_names, row))
                    result.append(row_dict)
            else:
                cursor = conn.cursor(cursor_factory=DictCursor)
                
                # Modified query to find the latest check-in/check-out status
                query = '''
                    SELECT a.*, h1.timestamp as check_out_date, h1.site
                    FROM assets a
                    JOIN (
                        SELECT asset_id, MAX(id) as max_id
                        FROM scan_history
                        WHERE status IN ('in', 'out')
                        GROUP BY asset_id
                    ) latest ON a.asset_id = latest.asset_id
                    JOIN scan_history h1 ON latest.max_id = h1.id
                    WHERE h1.status = 'out'
                '''
                
                # Add filter for deleted assets if requested
                if not include_deleted:
                    query += " AND (a.operational_status IS NULL OR a.operational_status != 'DELETED')"
                    
                query += " ORDER BY h1.timestamp DESC"
                
                cursor.execute(query)
                
                result = [dict(row) for row in cursor.fetchall()]
                    
            return result
        except Exception as e:
            logger.error(f"Error getting checked out inventory: {str(e)}")
            return []
        finally:
            if conn:
                self.release_connection(conn)
    
    def get_recent_history(self, days=30):
        """Get scan history for the last X days"""
        conn = None
        try:
            conn = self.get_connection(write=False)  # Read operation
            cursor = conn.cursor()
            
            if self.using_local:
                # SQLite date calculation is different
                date_limit = datetime.now().timestamp() - (days * 86400)  # days in seconds
                cursor.execute('''
                    SELECT h.id, h.asset_id, h.status, h.timestamp, h.notes, h.tech_name, a.serial_number, h.site
                    FROM scan_history h
                    JOIN assets a ON h.asset_id = a.asset_id
                    WHERE datetime(h.timestamp) >= datetime(?)
                    ORDER BY h.timestamp DESC
                ''', (date_limit,))
                
                # Get column names from cursor description
                column_names = [description[0] for description in cursor.description]
                
                # Convert rows to dictionaries
                result = []
                for row in cursor.fetchall():
                    row_dict = dict(zip(column_names, row))
                    result.append(row_dict)
            else:
                cursor = conn.cursor(cursor_factory=DictCursor)
                
                cursor.execute('''
                    SELECT h.*, a.serial_number
                    FROM scan_history h
                    JOIN assets a ON h.asset_id = a.asset_id
                    WHERE h.timestamp >= CURRENT_TIMESTAMP - INTERVAL %s DAY
                    ORDER BY h.timestamp DESC
                ''', (str(days),))
                
                result = [dict(row) for row in cursor.fetchall()]
                
            return result
        except Exception as e:
            logger.error(f"Error getting recent history: {str(e)}")
            return []
        finally:
            if conn:
                self.release_connection(conn)
    
    def search_asset_history(self, search_term):
        """Search full history for an asset"""
        conn = None
        try:
            conn = self.get_connection(write=False)  # Read operation
            cursor = conn.cursor()
            
            if self.using_local:
                cursor.execute('''
                    SELECT h.id, h.asset_id, h.status, h.timestamp, h.notes, h.tech_name, h.site
                    FROM scan_history h
                    JOIN assets a ON h.asset_id = a.asset_id
                    WHERE a.asset_id LIKE ? OR a.serial_number LIKE ?
                    ORDER BY h.timestamp DESC
                ''', (f'%{search_term}%', f'%{search_term}%'))
                
                # Get column names from cursor description
                column_names = [description[0] for description in cursor.description]
                
                # Convert rows to dictionaries
                result = []
                for row in cursor.fetchall():
                    row_dict = dict(zip(column_names, row))
                    result.append(row_dict)
            else:
                cursor = conn.cursor(cursor_factory=DictCursor)
                
                cursor.execute('''
                    SELECT h.*
                    FROM scan_history h
                    JOIN assets a ON h.asset_id = a.asset_id
                    WHERE a.asset_id LIKE %s OR a.serial_number LIKE %s
                    ORDER BY h.timestamp DESC
                ''', (f'%{search_term}%', f'%{search_term}%'))
                
                result = [dict(row) for row in cursor.fetchall()]
                
            return result
        except Exception as e:
            logger.error(f"Error searching asset history: {str(e)}")
            return []
        finally:
            if conn:
                self.release_connection(conn)
    
    def get_asset_by_id(self, asset_id):
        """Get a specific asset by ID"""
        conn = None
        try:
            conn = self.get_connection(write=False)  # Read operation
            cursor = conn.cursor()
            
            if self.using_local:
                cursor.execute("SELECT * FROM assets WHERE asset_id = ?", (asset_id,))
                
                # Get column names from cursor description
                column_names = [description[0] for description in cursor.description]
                
                # Fetch the row
                row = cursor.fetchone()
                
                # Convert row to dictionary
                result = dict(zip(column_names, row)) if row else None
            else:
                cursor = conn.cursor(cursor_factory=DictCursor)
                
                cursor.execute("SELECT * FROM assets WHERE asset_id = %s", (asset_id,))
                
                result = cursor.fetchone()
                if result:
                    result = dict(result)
                    
            return result
        except Exception as e:
            logger.error(f"Error getting asset by ID: {str(e)}")
            return None
        finally:
            if conn:
                self.release_connection(conn)
    
    def get_asset_by_serial(self, serial):
        """Get a specific asset by serial number"""
        conn = None
        try:
            conn = self.get_connection(write=False)  # Read operation
            cursor = conn.cursor()
            
            if self.using_local:
                cursor.execute("SELECT * FROM assets WHERE serial_number = ?", (serial,))
                
                # Get column names from cursor description
                column_names = [description[0] for description in cursor.description]
                
                # Fetch the row
                row = cursor.fetchone()
                
                # Convert row to dictionary
                result = dict(zip(column_names, row)) if row else None
            else:
                cursor = conn.cursor(cursor_factory=DictCursor)
                cursor.execute("SELECT * FROM assets WHERE serial_number = %s", (serial,))
                
                result = cursor.fetchone()
                if result:
                    result = dict(result)
                    
            return result
        except Exception as e:
            logger.error(f"Error getting asset by serial: {str(e)}")
            return None
        finally:
            if conn:
                self.release_connection(conn)
    
    def get_asset_history(self, asset_id, limit=5):
        """Get recent history for a specific asset"""
        conn = None
        try:
            conn = self.get_connection(write=False)  # Read operation
            cursor = conn.cursor()
            
            if self.using_local:
                cursor.execute('''
                    SELECT id, asset_id, status, timestamp, notes, tech_name, site
                    FROM scan_history
                    WHERE asset_id = ?
                    ORDER BY timestamp DESC
                    LIMIT ?
                ''', (asset_id, limit))
                
                # Get column names from cursor description
                column_names = [description[0] for description in cursor.description]
                
                # Convert rows to dictionaries
                result = []
                for row in cursor.fetchall():
                    row_dict = dict(zip(column_names, row))
                    result.append(row_dict)
            else:
                cursor = conn.cursor(cursor_factory=DictCursor)
                
                cursor.execute('''
                    SELECT * FROM scan_history
                    WHERE asset_id = %s
                    ORDER BY timestamp DESC
                    LIMIT %s
                ''', (asset_id, limit))
                
                result = [dict(row) for row in cursor.fetchall()]
                
            return result
        except Exception as e:
            logger.error(f"Error getting asset history: {str(e)}")
            return []
        finally:
            if conn:
                self.release_connection(conn)

    def delete_asset(self, asset_id):
        """Soft delete an asset (mark as deleted instead of removing)"""
        conn = None
        try:
            conn = self.get_connection(write=True)
            cursor = conn.cursor()
            
            # Update asset to mark as deleted
            if self.using_local:
                cursor.execute("""
                    UPDATE assets 
                    SET operational_status = 'DELETED', 
                        comments = CASE
                            WHEN comments IS NULL OR comments = '' THEN 'Asset deleted from inventory'
                            ELSE comments || ' | Asset deleted from inventory'
                        END
                    WHERE asset_id = ?
                """, (asset_id,))
            else:
                cursor.execute("""
                    UPDATE assets 
                    SET operational_status = 'DELETED', 
                        comments = CASE
                            WHEN comments IS NULL OR comments = '' THEN 'Asset deleted from inventory'
                            ELSE comments || ' | Asset deleted from inventory'
                        END
                    WHERE asset_id = %s
                """, (asset_id,))
            
            conn.commit()
            
            # Record operation for sync if using local DB
            if self.using_local:
                asset_data = {'asset_id': asset_id, 'operational_status': 'DELETED'}
                self._record_operation("UPDATE", "assets", asset_data)
            
            logger.info(f"Soft deleted asset {asset_id}")
            return True
        except Exception as e:
            logger.error(f"Error soft deleting asset: {str(e)}")
            if conn:
                conn.rollback()
            return False
        finally:
            if conn:
                self.release_connection(conn)
    
    def hard_delete_asset(self, asset_id):
        """Actually delete an asset from the database (for administrator use)"""
        conn = None
        try:
            conn = self.get_connection(write=True)
            cursor = conn.cursor()
            
            # Delete the asset
            if self.using_local:
                cursor.execute("DELETE FROM assets WHERE asset_id = ?", (asset_id,))
            else:
                cursor.execute("DELETE FROM assets WHERE asset_id = %s", (asset_id,))
            
            # Note: We keep the history records for auditing purposes
            
            conn.commit()
            
            # Record operation for sync if using local DB
            if self.using_local:
                asset_data = {'asset_id': asset_id}
                self._record_operation("DELETE", "assets", asset_data)
            
            logger.info(f"Hard deleted asset {asset_id}")
            return True
        except Exception as e:
            logger.error(f"Error hard deleting asset: {str(e)}")
            if conn:
                conn.rollback()
            return False
        finally:
            if conn:
                self.release_connection(conn)

    def cursor(self, conn):
        """Get appropriate cursor based on database type"""
        if self.using_local:
            # SQLite cursor
            return conn.cursor()
        else:
            # PostgreSQL cursor with dictionary support
            return conn.cursor(cursor_factory=DictCursor)

    def _sync_central_to_local(self):
        """
        Fetches data from the central PostgreSQL DB and overwrites the local SQLite DB.
        """
        logger.info("Starting sync from central PostgreSQL to local SQLite...")
        pg_conn = None
        sqlite_conn = None
        try:
            # Get connections
            pg_conn = self._get_pg_connection(write=False) # Read from primary/replica
            sqlite_conn = self._get_sqlite_connection()   # Write to local
            pg_cursor = pg_conn.cursor(cursor_factory=DictCursor) # Use DictCursor for PG
            sqlite_cursor = sqlite_conn.cursor()

            # === Sync Assets Table ===
            logger.info("Syncing assets table...")
            # 1. Fetch all assets from PostgreSQL
            pg_cursor.execute("SELECT * FROM assets")
            pg_assets = pg_cursor.fetchall()

            # 2. Clear local assets table
            sqlite_cursor.execute("DELETE FROM assets")
            logger.info(f"Cleared local assets table.")

            # 3. Insert fetched assets into SQLite
            if pg_assets:
                asset_columns = list(pg_assets[0].keys())
                # Ensure flag_status is converted for SQLite (0/1)
                sqlite_asset_columns = [col if col != 'flag_status' else 'flag_status' for col in asset_columns]
                placeholders = ', '.join(['?'] * len(sqlite_asset_columns))
                insert_sql = f"INSERT INTO assets ({', '.join(sqlite_asset_columns)}) VALUES ({placeholders})"

                sqlite_rows = []
                for asset in pg_assets:
                    row_values = []
                    for col in asset_columns:
                        value = asset[col]
                        if col == 'flag_status':
                            row_values.append(1 if value else 0) # Convert boolean to integer
                        elif isinstance(value, datetime):
                            # Store timestamps as ISO format strings
                            row_values.append(value.isoformat())
                        else:
                            row_values.append(value)
                    sqlite_rows.append(tuple(row_values))

                sqlite_cursor.executemany(insert_sql, sqlite_rows)
                logger.info(f"Inserted {len(pg_assets)} assets into local DB.")

            # === Sync Scan History Table (Optional - decide scope) ===
            # Syncing full history can be large. Syncing recent history might be better.
            # Example: Sync last 90 days of history
            days_to_sync = 90
            logger.info(f"Syncing scan_history table (last {days_to_sync} days)...")
            # 1. Fetch recent history from PostgreSQL
            pg_cursor.execute("""
                SELECT * FROM scan_history
                WHERE timestamp >= CURRENT_TIMESTAMP - INTERVAL '%s days'
            """, (days_to_sync,))
            pg_history = pg_cursor.fetchall()

            # 2. Clear local scan_history table (or selectively delete old records)
            sqlite_cursor.execute("DELETE FROM scan_history") # Simple clear for this example
            logger.info(f"Cleared local scan_history table.")

            # 3. Insert fetched history into SQLite
            if pg_history:
                history_columns = list(pg_history[0].keys())
                history_placeholders = ', '.join(['?'] * len(history_columns))
                history_insert_sql = f"INSERT INTO scan_history ({', '.join(history_columns)}) VALUES ({history_placeholders})"

                sqlite_history_rows = []
                for record in pg_history:
                    row_values = []
                    for col in history_columns:
                        value = record[col]
                        if isinstance(value, datetime):
                            row_values.append(value.isoformat())
                        else:
                            row_values.append(value)
                    sqlite_history_rows.append(tuple(row_values))

                sqlite_cursor.executemany(history_insert_sql, sqlite_history_rows)
                logger.info(f"Inserted {len(pg_history)} recent history records into local DB.")

            # === Sync Related Items Table (If needed) ===
            # Add similar logic for related_items if required

            # Commit changes to local SQLite DB
            sqlite_conn.commit()
            logger.info("Sync from central to local successfully committed.")

        except Exception as e:
            logger.error(f"Error during central to local sync: {e}", exc_info=True)
            if sqlite_conn:
                sqlite_conn.rollback() # Rollback local changes on error
            # Re-raise the exception to be caught in __init__
            raise
        finally:
            if pg_conn:
                self.release_connection(pg_conn)
            if sqlite_conn:
                sqlite_conn.close()

    def _sync_to_server(self):
        """Sync local changes from sync_queue (SQLite) to PostgreSQL server."""
        # Sync only makes sense if there were potentially pending changes
        # No need to run if pending_sync is already False
        if not self.pending_sync:
            logger.info("Sync to server: No pending changes flagged.")
            return

        logger.info("Sync to server: Starting synchronization...")
        pg_conn = None
        sqlite_conn = None # Connection specifically for reading/writing sync_queue
        processed_ids = set()
        pending_changes = []

        # --- Step 1: Read pending changes directly from SQLite ---
        try:
            logger.info(f"Sync: Connecting directly to SQLite DB ({self.local_db_path}) to read sync_queue...")
            # Ensure we use the correct path, potentially use shared connection if robustly handled,
            # but a separate connection for this specific task is also fine.
            # Let's use a separate connection here for clarity.
            sqlite_conn = sqlite3.connect(self.local_db_path)
            sqlite_cursor = sqlite_conn.cursor()
            sqlite_cursor.execute("SELECT id, operation, table_name, data, timestamp FROM sync_queue ORDER BY timestamp")
            pending_changes = sqlite_cursor.fetchall()

            if not pending_changes:
                self.pending_sync = False
                logger.info("Sync: No pending changes found in local sync queue.")
                sqlite_conn.close() # Close the connection used for reading
                return

            logger.info(f"Sync: Found {len(pending_changes)} pending changes in local sync queue.")

        except Exception as read_err:
            logger.error(f"Sync: Failed to read sync_queue from SQLite ({self.local_db_path}): {read_err}", exc_info=True)
            if sqlite_conn:
                 try: sqlite_conn.close()
                 except Exception: pass
            # Can't proceed without reading the queue, keep pending_sync=True
            self.pending_sync = True
            return # Exit sync process

        # --- Step 2: Connect to PostgreSQL to apply changes ---
        try:
            # Make sure pools are initialized (might be redundant if checker worked, but safe)
            if not self.primary_pool: self._initialize_pg_connection_pools()
            pg_conn = self._get_pg_connection(write=True) # Now get the PG connection
            pg_cursor = pg_conn.cursor()
            logger.info("Sync: Connected to PostgreSQL to apply changes.")
        except Exception as e:
            logger.error(f"Sync: Failed to connect to PostgreSQL to apply changes: {e}")
            if sqlite_conn: # Close the SQLite reader connection if still open
                 try: sqlite_conn.close()
                 except Exception: pass
            # Can't sync, keep pending_sync=True
            self.pending_sync = True
            return

        # --- Step 3: Apply each change to PostgreSQL ---
        # (The loop for applying changes remains largely the same as the previous version)
        for change in pending_changes:
            change_id, operation, table, data_json, timestamp = change
            data = None

            try:
                data = json.loads(data_json)

                # Type conversion for flag_status
                if table == "assets" and 'flag_status' in data:
                    try:
                        original_value = data['flag_status']
                        data['flag_status'] = bool(original_value)
                        logger.debug(f"Converted flag_status from {original_value} to {data['flag_status']} for sync ID {change_id}")
                    except Exception as conv_err:
                        raise ValueError(f"Invalid flag_status value '{data['flag_status']}': {conv_err}")

                # Apply INSERT, UPDATE, DELETE using pg_cursor
                if operation == "INSERT":
                    columns = ", ".join(data.keys())
                    placeholders = ", ".join(["%s"] * len(data))
                    query = f"INSERT INTO {table} ({columns}) VALUES ({placeholders})"
                    pg_cursor.execute(query, list(data.values()))
                    logger.info(f"Sync: Applied INSERT for ID {change_id} to PG table {table}")

                elif operation == "UPDATE":
                    id_field = "asset_id" if table == "assets" else "id"
                    if id_field not in data: raise KeyError(f"Missing primary key '{id_field}' in update data for sync ID {change_id}")
                    id_value = data[id_field]
                    update_data = {k: v for k, v in data.items() if k != id_field}
                    if not update_data: raise ValueError(f"No fields to update for sync ID {change_id}")

                    set_clause = ", ".join([f"{k} = %s" for k in update_data.keys()])
                    values = list(update_data.values()) + [id_value]
                    query = f"UPDATE {table} SET {set_clause} WHERE {id_field} = %s"
                    pg_cursor.execute(query, values)
                    logger.info(f"Sync: Applied UPDATE for ID {change_id} to PG table {table}")

                elif operation == "DELETE":
                    id_field = "asset_id" if table == "assets" else "id"
                    if id_field not in data: raise KeyError(f"Missing primary key '{id_field}' in delete data for sync ID {change_id}")
                    id_value = data[id_field]
                    query = f"DELETE FROM {table} WHERE {id_field} = %s"
                    pg_cursor.execute(query, [id_value])
                    logger.info(f"Sync: Applied DELETE for ID {change_id} to PG table {table}")

                # Mark change for removal from SQLite queue if PG operation succeeded
                processed_ids.add(change_id)
                pg_conn.commit() # Commit this single successful PG operation

            except (psycopg2.Error, json.JSONDecodeError, KeyError, ValueError) as op_err:
                logger.error(f"Sync: Error applying operation to PG for sync ID {change_id} ({operation} on {table}): {op_err}")
                logger.error(f"Sync: Failing data snippet: {data_json[:200]}...")
                if pg_conn:
                     try:
                          pg_conn.rollback() # Rollback the failed PG operation
                     except psycopg2.Error as rb_err:
                          logger.error(f"Sync: Error during PG rollback: {rb_err}")
                # Do NOT add change_id to processed_ids, leave it in the queue
                # Continue to the next change in the loop

        # --- Step 4: Delete successfully processed items from SQLite queue ---
        if processed_ids:
             try:
                  # Ensure the SQLite connection used for reading is still available
                  if sqlite_conn is None or sqlite_conn.total_changes == -1: # Check if closed
                       sqlite_conn = sqlite3.connect(self.local_db_path)
                       sqlite_cursor = sqlite_conn.cursor()

                  placeholders = ','.join('?' * len(processed_ids))
                  sqlite_cursor.execute(f"DELETE FROM sync_queue WHERE id IN ({placeholders})", tuple(processed_ids))
                  sqlite_conn.commit() # Commit deletions from SQLite queue
                  logger.info(f"Sync: Removed {len(processed_ids)} successfully processed items from local sync queue.")
             except Exception as del_err:
                  logger.error(f"Sync: Error deleting processed items from sync_queue: {del_err}")
                  if sqlite_conn:
                       try: sqlite_conn.rollback()
                       except Exception: pass

        # --- Step 5: Final check if queue is empty ---
        try:
             if sqlite_conn is None or sqlite_conn.total_changes == -1: # Check if closed
                  sqlite_conn = sqlite3.connect(self.local_db_path)
                  sqlite_cursor = sqlite_conn.cursor()
             sqlite_cursor.execute("SELECT COUNT(*) FROM sync_queue")
             remaining_count = sqlite_cursor.fetchone()[0]
             if remaining_count == 0:
                  self.pending_sync = False
                  logger.info("Sync to server: Synchronization completed. No items remaining in local queue.")
             else:
                  self.pending_sync = True
                  logger.warning(f"Sync to server: Synchronization finished, but {remaining_count} items failed and remain in the local queue.")
        except Exception as count_err:
             logger.error(f"Sync: Error checking remaining sync queue count: {count_err}")
             self.pending_sync = True # Assume items remain

        # --- Step 6: Clean up connections ---
        finally:
            if pg_conn:
                self.release_connection(pg_conn) # Release PG connection back to pool
            if sqlite_conn:
                 try:
                      sqlite_conn.close() # Close the connection used for sync queue operations
                 except Exception as close_err:
                      logger.error(f"Sync: Error closing SQLite connection: {close_err}")

    def _start_connectivity_checker(self):
        """Start a background thread to check server connectivity"""
        import threading
        
        def check_connectivity():
            while True:
                try:
                    # Sleep first
                    time.sleep(30) # Check every 30 seconds

                    if self.using_local: # Only try if currently offline
                         is_connected = False
                         # --- FIX: Try to re-initialize pools first ---
                         logger.debug("Connectivity checker: Currently local, attempting reconnect...")
                         try:
                             logger.info("Connectivity checker: Attempting pool re-initialization...")
                             self._initialize_pg_connection_pools()
                             logger.info("Connectivity checker: Pools re-initialized.")

                             # Test connection briefly
                             conn_test = self._get_pg_connection(write=False)
                             self.release_connection(conn_test)
                             is_connected = True
                             logger.info("Connectivity checker: Test connection successful.")

                         except Exception as check_err:
                             is_connected = False
                             # Log the specific error during the check if needed
                             logger.debug(f"Connectivity checker: Reconnect attempt failed: {check_err}")
                             # Make sure pools are None if init failed
                             if self.primary_pool: self.primary_pool.closeall(); self.primary_pool = None
                             self.replica_pools = [] # Assuming replicas are less critical for basic check

                         # --- END FIX ---

                         if is_connected:
                             logger.info("Reconnected to PostgreSQL server automatically.")
                             self.using_local = False # Switch back to online mode

                             # Trigger sync only if there were pending changes
                             if self.pending_sync:
                                 logger.info("Attempting automatic background sync...")
                                 self._sync_to_server()
                             # Update last sync time potentially
                             self.last_sync_time = datetime.now()

                             # No need to explicitly update UI indicator here,
                             # the main UI thread's update_status_indicator will pick it up.

                except Exception as e:
                    logger.error(f"Error in connectivity checker thread: {e}", exc_info=True)
        
        # Start checker thread
        threading.Thread(target=check_connectivity, daemon=True).start()

    def _record_operation(self, operation, table, data):
        """Record operation in sync queue when using local database"""
        if not self.using_local:
            return

        conn = None # Remove conn initialization here
        try: # Wrap the operation
            conn = self.get_connection(write=True) # Get shared connection
            cursor = conn.cursor()

            # Convert data to JSON (ensure datetimes are handled before this call)
            data_json = json.dumps(data)

            # Add to sync queue
            cursor.execute(
                "INSERT INTO sync_queue (operation, table_name, data, timestamp) VALUES (?, ?, ?, ?)",
                (operation, table, data_json, datetime.now().isoformat())
            )
            conn.commit()
            # conn.close() # <<< REMOVE this line
            self.pending_sync = True
        except Exception as e:
             logger.error(f"Error in _record_operation: {e}", exc_info=True)
             if conn:
                  try:
                       conn.rollback() # Attempt rollback on error
                  except Exception as rb_err:
                       logger.error(f"Error during rollback in _record_operation: {rb_err}")

    def update_lease_info(self, asset_id, lease_start_date=None, lease_maturity_date=None):
        """Update lease information for a specific asset"""
        conn = None
        try:
            conn = self.get_connection(write=True)
            cursor = conn.cursor()
            
            update_data = {}
            if lease_start_date is not None:
                update_data['lease_start_date'] = lease_start_date
            if lease_maturity_date is not None:
                update_data['lease_maturity_date'] = lease_maturity_date
            
            if not update_data:
                logger.warning(f"No lease data provided for asset: {asset_id}")
                return True  # Nothing to update
            
            # Set last updated timestamp
            update_data['last_updated'] = datetime.now()
            
            # Update asset
            set_clause = ", ".join([f"{key} = {'?' if self.using_local else '%s'}" for key in update_data.keys()])
            values = list(update_data.values())
            values.append(asset_id)
            
            query = f"UPDATE assets SET {set_clause} WHERE asset_id = {'?' if self.using_local else '%s'}"
            cursor.execute(query, values)
            
            conn.commit()
            logger.info(f"Updated lease information for asset {asset_id}")
            
            # Record operation for sync if using local DB
            if self.using_local:
                update_data['asset_id'] = asset_id
                self._record_operation("UPDATE", "assets", update_data)
            
            # Check if we need to update expiry flag status based on lease maturity date
            if lease_maturity_date:
                self.check_and_update_expiry_flag(asset_id)
            
            return True
        except Exception as e:
            logger.error(f"Database error updating lease info: {str(e)}")
            if conn:
                conn.rollback()
            return False
        finally:
            if conn:
                self.release_connection(conn)

    def check_and_update_expiry_flag(self, asset_id):
        """Check lease maturity date and update expiry flag if within 90 days or already expired"""
        conn = None
        try:
            conn = self.get_connection(write=True)
            cursor = conn.cursor()
            
            # Get the asset's lease maturity date
            if self.using_local:
                cursor.execute("SELECT lease_maturity_date FROM assets WHERE asset_id = ?", (asset_id,))
            else:
                cursor.execute("SELECT lease_maturity_date FROM assets WHERE asset_id = %s", (asset_id,))
            
            result = cursor.fetchone()
            if not result or not result[0]:
                # No maturity date, nothing to do
                return True
            
            lease_maturity_date = result[0]
            
            # Parse the date from string
            try:
                if isinstance(lease_maturity_date, str):
                    # Handle different date formats
                    for fmt in ['%Y-%m-%d', '%m/%d/%Y', '%d/%m/%Y', '%Y/%m/%d']:
                        try:
                            maturity_date = datetime.strptime(lease_maturity_date, fmt).date()
                            break
                        except ValueError:
                            continue
                    else:
                        # If no format matched
                        logger.warning(f"Unrecognized date format for asset {asset_id}: {lease_maturity_date}")
                        return False
                elif isinstance(lease_maturity_date, datetime):
                    maturity_date = lease_maturity_date.date()
                else:
                    logger.warning(f"Unexpected date type for asset {asset_id}: {type(lease_maturity_date)}")
                    return False
            except Exception as e:
                logger.error(f"Error parsing lease maturity date for asset {asset_id}: {e}")
                return False
            
            # Check if within 90 days or already expired
            today = datetime.now().date()
            days_remaining = (maturity_date - today).days
            
            # Flag if less than or equal to 90 days remaining, OR already expired (negative days_remaining)
            should_flag = days_remaining <= 90
            
            # Update expiry flag status
            if self.using_local:
                cursor.execute("""
                    UPDATE assets 
                    SET expiry_flag_status = ?, last_updated = ?
                    WHERE asset_id = ?
                """, (1 if should_flag else 0, datetime.now(), asset_id))
            else:
                cursor.execute("""
                    UPDATE assets 
                    SET expiry_flag_status = %s, last_updated = %s
                    WHERE asset_id = %s
                """, (True if should_flag else False, datetime.now(), asset_id))
            
            conn.commit()
            
            # Record operation for sync if using local DB
            if self.using_local:
                update_data = {
                    'asset_id': asset_id,
                    'expiry_flag_status': 1 if should_flag else 0,
                    'last_updated': datetime.now().isoformat()
                }
                self._record_operation("UPDATE", "assets", update_data)
            
            logger.info(f"Updated expiry flag for asset {asset_id} to {should_flag} (days remaining: {days_remaining})")
            return True
        except Exception as e:
            logger.error(f"Error checking expiry flag for asset {asset_id}: {str(e)}")
            if conn:
                conn.rollback()
            return False
        finally:
            if conn:
                self.release_connection(conn)

    def get_expiry_flag_status(self, asset_id):
        """Get the expiry flag status and lease details for an asset"""
        conn = None
        try:
            conn = self.get_connection(write=False)  # Read operation
            cursor = conn.cursor()
            
            if self.using_local:
                cursor.execute('''
                    SELECT expiry_flag_status, lease_start_date, lease_maturity_date
                    FROM assets
                    WHERE asset_id = ?
                ''', (asset_id,))
                
                # Get column names from cursor description
                column_names = ["expiry_flag_status", "lease_start_date", "lease_maturity_date"]
                
                # Fetch the row
                row = cursor.fetchone()
                
                # Convert row to dictionary
                if row:
                    result = dict(zip(column_names, row))
                    # Convert expiry_flag_status from integer to boolean for consistency
                    result['expiry_flag_status'] = bool(result['expiry_flag_status'])
                    return result
                else:
                    return {'expiry_flag_status': False, 'lease_start_date': None, 'lease_maturity_date': None}
            else:
                cursor = conn.cursor(cursor_factory=DictCursor)
                
                cursor.execute('''
                    SELECT expiry_flag_status, lease_start_date, lease_maturity_date
                    FROM assets
                    WHERE asset_id = %s
                ''', (asset_id,))
                
                result = cursor.fetchone()
                if result:
                    return dict(result)
                else:
                    return {'expiry_flag_status': False, 'lease_start_date': None, 'lease_maturity_date': None}
                    
        except Exception as e:
            logger.error(f"Error getting expiry flag status: {str(e)}")
            return {'expiry_flag_status': False, 'lease_start_date': None, 'lease_maturity_date': None}
        finally:
            if conn:
                self.release_connection(conn)

    def get_expiring_assets(self, days=90, include_deleted=False):
        """Get all assets that are expiring within the specified number of days"""
        conn = None
        try:
            conn = self.get_connection(write=False)  # Read operation
            cursor = conn.cursor()
            
            today = datetime.now().date()
            
            # We'll filter the results in Python since date handling in SQLite 
            # and PostgreSQL can be different, and we need to support different date formats
            if self.using_local:
                if include_deleted:
                    cursor.execute('''
                        SELECT * FROM assets 
                        WHERE expiry_flag_status = 1
                    ''')
                else:
                    cursor.execute('''
                        SELECT * FROM assets 
                        WHERE expiry_flag_status = 1
                        AND (operational_status IS NULL OR operational_status != 'DELETED')
                    ''')
                    
                # Get column names from cursor description
                column_names = [description[0] for description in cursor.description]
                
                # Fetch all rows
                rows = cursor.fetchall()
                
                # Convert rows to dictionaries
                assets = []
                for row in rows:
                    asset = dict(zip(column_names, row))
                    
                    # Process the lease maturity date
                    maturity_date = asset.get('lease_maturity_date')
                    if not maturity_date:
                        continue
                    
                    # Add to results
                    assets.append(asset)
                    
            else:
                cursor = conn.cursor(cursor_factory=DictCursor)
                
                if include_deleted:
                    cursor.execute('''
                        SELECT * FROM assets 
                        WHERE expiry_flag_status = TRUE
                    ''')
                else:
                    cursor.execute('''
                        SELECT * FROM assets 
                        WHERE expiry_flag_status = TRUE
                        AND (operational_status IS NULL OR operational_status != 'DELETED')
                    ''')
                
                assets = [dict(row) for row in cursor.fetchall()]
            
            # Get current status for each asset
            for asset in assets:
                status_info = self.get_asset_current_status(asset['asset_id'])
                if status_info:
                    asset['status'] = status_info.get('status')
                    asset['site'] = status_info.get('site')
            
            return assets
        except Exception as e:
            logger.error(f"Error getting expiring assets: {str(e)}")
            return []
        finally:
            if conn:
                self.release_connection(conn)

    def process_lease_data_from_file(self, file_path):
        """Process lease data from an Excel file and update the database"""
        try:
            import pandas as pd
            
            # Read the Excel file
            if file_path.lower().endswith('.xlsx'):
                df = pd.read_excel(file_path)
            elif file_path.lower().endswith('.csv'):
                df = pd.read_csv(file_path)
            else:
                logger.error(f"Unsupported file format: {file_path}")
                return False, "Unsupported file format. Please use .xlsx or .csv"
            
            # Check required columns
            required_columns = ['Serial Number', 'Lease Start Date', 'Lease Maturity Date']
            if not all(col in df.columns for col in required_columns):
                logger.error(f"Missing required columns. Required: {required_columns}, Found: {list(df.columns)}")
                return False, f"Missing required columns. Required: {required_columns}"
            
            # Process each row
            total_rows = len(df)
            updated_count = 0
            not_found_count = 0
            error_count = 0
            not_found_serials = []
            
            for _, row in df.iterrows():
                serial_number = str(row['Serial Number']).strip()
                if not serial_number or pd.isna(serial_number):
                    continue
                    
                # Parse dates
                lease_start_date = None
                lease_maturity_date = None
                
                if 'Lease Start Date' in row and not pd.isna(row['Lease Start Date']):
                    start_date = row['Lease Start Date']
                    if isinstance(start_date, str):
                        lease_start_date = start_date
                    else:
                        # Convert to string in YYYY-MM-DD format
                        lease_start_date = pd.to_datetime(start_date).strftime('%Y-%m-%d')
                
                if 'Lease Maturity Date' in row and not pd.isna(row['Lease Maturity Date']):
                    maturity_date = row['Lease Maturity Date']
                    if isinstance(maturity_date, str):
                        lease_maturity_date = maturity_date
                    else:
                        # Convert to string in YYYY-MM-DD format
                        lease_maturity_date = pd.to_datetime(maturity_date).strftime('%Y-%m-%d')
                
                # Skip if both dates are None
                if lease_start_date is None and lease_maturity_date is None:
                    continue
                
                # Find asset by serial number
                asset = self.get_asset_by_serial(serial_number)
                if not asset:
                    not_found_count += 1
                    not_found_serials.append(serial_number)
                    continue
                
                # Update lease info
                try:
                    success = self.update_lease_info(
                        asset['asset_id'],
                        lease_start_date=lease_start_date,
                        lease_maturity_date=lease_maturity_date
                    )
                    
                    if success:
                        updated_count += 1
                    else:
                        error_count += 1
                except Exception as e:
                    logger.error(f"Error updating lease info for {serial_number}: {str(e)}")
                    error_count += 1
            
            # Log results
            logger.info(f"Processed {total_rows} rows from {file_path}")
            logger.info(f"Updated: {updated_count}, Not found: {not_found_count}, Errors: {error_count}")
            
            if not_found_count > 0:
                logger.info(f"Serials not found: {not_found_serials[:10]}...")
            
            return True, f"Processed {total_rows} rows. Updated: {updated_count}, Not found: {not_found_count}, Errors: {error_count}"
            
        except Exception as e:
            logger.error(f"Error processing lease data file: {str(e)}")
            return False, f"Error processing file: {str(e)}"

    def update_all_expiry_flags(self):
        """Check all assets with lease maturity dates and update their expiry flags"""
        conn = None
        try:
            conn = self.get_connection(write=False)  # Read operation first to get assets
            cursor = conn.cursor()
            
            # Get all assets with lease maturity dates
            if self.using_local:
                cursor.execute('''
                    SELECT asset_id, lease_maturity_date 
                    FROM assets 
                    WHERE lease_maturity_date IS NOT NULL 
                    AND lease_maturity_date != ''
                ''')
            else:
                cursor.execute('''
                    SELECT asset_id, lease_maturity_date 
                    FROM assets 
                    WHERE lease_maturity_date IS NOT NULL 
                    AND lease_maturity_date != ''
                ''')
            
            assets_to_check = cursor.fetchall()
            
            # Release the connection since we're done reading
            self.release_connection(conn)
            conn = None
            
            updated_count = 0
            error_count = 0
            
            # Check each asset's expiry flag
            for asset_info in assets_to_check:
                asset_id = asset_info[0]
                try:
                    success = self.check_and_update_expiry_flag(asset_id)
                    if success:
                        updated_count += 1
                    else:
                        error_count += 1
                except Exception as e:
                    logger.error(f"Error updating expiry flag for {asset_id}: {str(e)}")
                    error_count += 1
            
            logger.info(f"Updated expiry flags for {updated_count} assets. Errors: {error_count}")
            return True
        except Exception as e:
            logger.error(f"Error updating all expiry flags: {str(e)}")
            return False
        finally:
            if conn:
                self.release_connection(conn)

    def flag_asset(self, asset_id, flag_notes, flag_tech):
        """Flag an asset as needing attention and automatically check it out"""
        conn = None
        try:
            conn = self.get_connection(write=True)
            cursor = conn.cursor()
            
            # First, check the current status
            current_status = self.get_asset_current_status(asset_id)
            current_state = current_status.get('status', 'unknown') if current_status else 'unknown'
            
            # Set the flag status, notes, timestamp, and tech
            timestamp = datetime.now()
            
            if self.using_local:
                cursor.execute('''
                    UPDATE assets 
                    SET flag_status = 1, 
                        flag_notes = ?, 
                        flag_timestamp = ?, 
                        flag_tech = ?
                    WHERE asset_id = ?
                ''', (flag_notes, timestamp, flag_tech, asset_id))
            else:
                cursor.execute('''
                    UPDATE assets 
                    SET flag_status = TRUE, 
                        flag_notes = %s, 
                        flag_timestamp = %s, 
                        flag_tech = %s
                    WHERE asset_id = %s
                ''', (flag_notes, timestamp, flag_tech, asset_id))
            
            conn.commit()
            
            # Record operation for sync if using local DB
            if self.using_local:
                flag_data = {
                    'asset_id': asset_id,
                    'flag_status': 1,
                    'flag_notes': flag_notes,
                    'flag_timestamp': timestamp.isoformat(),
                    'flag_tech': flag_tech
                }
                self._record_operation("UPDATE", "assets", flag_data)
                
            logger.info(f"Flagged asset {asset_id}")
            
            # Record this action in the scan history with site set to "Out"
            self.record_scan(asset_id, "flagged", flag_tech, flag_notes, site="Out")
            
            # If the asset is currently checked in, check it out automatically
            if current_state == "in":
                logger.info(f"Automatically checking out flagged asset {asset_id}")
                # Record a check-out action with a note about auto-checkout due to flagging
                checkout_note = f"Automatically checked out due to flagging: {flag_notes}"
                self.record_scan(asset_id, "out", flag_tech, checkout_note, site="Out")
            
            return True
        except Exception as e:
            logger.error(f"Error flagging asset: {str(e)}")
            if conn:
                conn.rollback()
            return False
        finally:
            if conn:
                self.release_connection(conn)

    def unflag_asset(self, asset_id, tech_name, notes=""):
        """Remove the flag from an asset"""
        conn = None
        try:
            conn = self.get_connection(write=True)
            cursor = conn.cursor()
            
            # First, get the current status to preserve the site information
            current_status = self.get_asset_current_status(asset_id)
            current_site = current_status.get('site') if current_status else None
            
            if self.using_local:
                cursor.execute('''
                    UPDATE assets 
                    SET flag_status = 0, 
                        flag_notes = NULL, 
                        flag_timestamp = NULL, 
                        flag_tech = NULL
                    WHERE asset_id = ?
                ''', (asset_id,))
            else:
                cursor.execute('''
                    UPDATE assets 
                    SET flag_status = FALSE, 
                        flag_notes = NULL, 
                        flag_timestamp = NULL, 
                        flag_tech = NULL
                    WHERE asset_id = %s
                ''', (asset_id,))
            
            conn.commit()
            
            # Record operation for sync if using local DB
            if self.using_local:
                flag_data = {
                    'asset_id': asset_id,
                    'flag_status': 0,
                    'flag_notes': None,
                    'flag_timestamp': None,
                    'flag_tech': None
                }
                self._record_operation("UPDATE", "assets", flag_data)
                    
            logger.info(f"Unflagged asset {asset_id}")
            
            # Record this action in the scan history - preserve the current site
            self.record_scan(asset_id, "unflagged", tech_name, notes, site=current_site)
            
            return True
        except Exception as e:
            logger.error(f"Error unflagging asset: {str(e)}")
            if conn:
                conn.rollback()
            return False
        finally:
            if conn:
                self.release_connection(conn)

    def get_flag_status(self, asset_id):
        """Get the flag status and details for an asset"""
        conn = None
        try:
            conn = self.get_connection(write=False)  # Read operation
            cursor = conn.cursor()
            
            if self.using_local:
                cursor.execute('''
                    SELECT flag_status, flag_notes, flag_timestamp, flag_tech
                    FROM assets
                    WHERE asset_id = ?
                ''', (asset_id,))
                
                # Get column names from cursor description
                column_names = ["flag_status", "flag_notes", "flag_timestamp", "flag_tech"]
                
                # Fetch the row
                row = cursor.fetchone()
                
                # Convert row to dictionary
                if row:
                    result = dict(zip(column_names, row))
                    # Convert flag_status from integer to boolean for consistency
                    result['flag_status'] = bool(result['flag_status'])
                    return result
                else:
                    return {'flag_status': False, 'flag_notes': None, 'flag_timestamp': None, 'flag_tech': None}
            else:
                cursor = conn.cursor(cursor_factory=DictCursor)
                
                cursor.execute('''
                    SELECT flag_status, flag_notes, flag_timestamp, flag_tech
                    FROM assets
                    WHERE asset_id = %s
                ''', (asset_id,))
                
                result = cursor.fetchone()
                if result:
                    return dict(result)
                else:
                    return {'flag_status': False, 'flag_notes': None, 'flag_timestamp': None, 'flag_tech': None}
                    
        except Exception as e:
            logger.error(f"Error getting flag status: {str(e)}")
            return {'flag_status': False, 'flag_notes': None, 'flag_timestamp': None, 'flag_tech': None}
        finally:
            if conn:
                self.release_connection(conn)

    def get_flagged_assets(self):
        """Get all assets that are currently flagged"""
        conn = None
        try:
            conn = self.get_connection(write=False)  # Read operation
            cursor = conn.cursor()
            
            if self.using_local:
                cursor.execute('''
                    SELECT a.*, h.status, h.timestamp, h.site
                    FROM assets a
                    LEFT JOIN (
                        SELECT asset_id, status, timestamp, site, 
                            ROW_NUMBER() OVER (PARTITION BY asset_id ORDER BY timestamp DESC) as rn
                        FROM scan_history
                        WHERE status IN ('in', 'out')
                    ) h ON a.asset_id = h.asset_id AND h.rn = 1
                    WHERE a.flag_status = 1
                    ORDER BY a.flag_timestamp DESC
                ''')
                
                # Get column names from cursor description
                column_names = [description[0] for description in cursor.description]
                
                # Convert rows to dictionaries
                result = []
                for row in cursor.fetchall():
                    row_dict = dict(zip(column_names, row))
                    result.append(row_dict)
            else:
                cursor = conn.cursor(cursor_factory=DictCursor)
                
                cursor.execute('''
                    SELECT a.*, h.status, h.timestamp, h.site
                    FROM assets a
                    LEFT JOIN (
                        SELECT asset_id, status, timestamp, site, 
                            ROW_NUMBER() OVER (PARTITION BY asset_id ORDER BY timestamp DESC) as rn
                        FROM scan_history
                        WHERE status IN ('in', 'out')
                    ) h ON a.asset_id = h.asset_id AND h.rn = 1
                    WHERE a.flag_status = TRUE
                    ORDER BY a.flag_timestamp DESC
                ''')
                
                result = [dict(row) for row in cursor.fetchall()]
                    
            return result
        except Exception as e:
            logger.error(f"Error getting flagged assets: {str(e)}")
            return []
        finally:
            if conn:
                self.release_connection(conn)

    def get_pending_changes_count(self):
        """Get the number of pending changes to be synced"""
        if not self.using_local or not self.pending_sync:
            return 0
            
        conn = None
        try:
            conn = self._get_sqlite_connection()
            cursor = conn.cursor()
            cursor.execute("SELECT COUNT(*) FROM sync_queue")
            count = cursor.fetchone()[0]
            return count
        except Exception as e:
            logger.error(f"Error getting pending changes count: {e}")
            return 0

    def close_db(self):
        logger.info("Closing database connections...")
        # Close shared SQLite connection if it exists
        if self.sqlite_conn:
            try:
                logger.info("Closing shared SQLite connection.")
                self.sqlite_conn.close()
                self.sqlite_conn = None
            except Exception as e:
                logger.error(f"Error closing SQLite connection: {e}")

        # Close PostgreSQL pools
        if self.primary_pool:
            try:
                logger.info("Closing primary PostgreSQL pool.")
                self.primary_pool.closeall()
                self.primary_pool = None
            except Exception as e:
                logger.error(f"Error closing primary PG pool: {e}")
        for i, pool_obj in enumerate(self.replica_pools):
            if pool_obj:
                try:
                    logger.info(f"Closing replica PostgreSQL pool {i}.")
                    pool_obj.closeall()
                except Exception as e:
                    logger.error(f"Error closing replica PG pool {i}: {e}")
        self.replica_pools = []
        logger.info("Database connections closed.")