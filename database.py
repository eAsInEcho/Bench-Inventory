import sqlite3
import os
from datetime import datetime
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

class InventoryDatabase:
    def __init__(self, db_path="inventory.db"):
        self.db_path = db_path
        self.initialize_database()
        
    def initialize_database(self):
        """Create the database if it doesn't exist"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Create tables if they don't exist
        cursor.executescript('''
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
                last_updated DATETIME
            );
            
            CREATE TABLE IF NOT EXISTS scan_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                asset_id TEXT,
                status TEXT,
                timestamp DATETIME DEFAULT CURRENT_TIMESTAMP,
                notes TEXT,
                tech_name TEXT,
                FOREIGN KEY (asset_id) REFERENCES assets(asset_id)
            );
            
            CREATE TABLE IF NOT EXISTS related_items (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                parent_asset_id TEXT,
                serial_number TEXT UNIQUE,
                item_type TEXT,
                notes TEXT,
                FOREIGN KEY (parent_asset_id) REFERENCES assets(asset_id)
            );
        ''')
        
        conn.commit()
        conn.close()
        logger.info(f"Database initialized at {self.db_path}")
    
    def update_asset(self, asset_data):
        """Update or insert asset in database"""
        if not asset_data:
            logger.warning("Cannot update: No asset data provided")
            return False
            
        # Make sure we have an asset_tag
        asset_tag = asset_data.get('asset_tag')
        if not asset_tag:
            logger.warning("Cannot update: No asset_tag provided")
            return False
            
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        # Check if asset exists
        cursor.execute("SELECT asset_id FROM assets WHERE asset_id = ?", (asset_tag,))
        exists = cursor.fetchone()
        
        # Set last updated timestamp
        asset_data['last_updated'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        try:
            if exists:
                # Update existing asset
                set_clause = ", ".join([f"{key} = ?" for key in asset_data.keys() if key != 'asset_tag'])
                values = [asset_data[key] for key in asset_data.keys() if key != 'asset_tag']
                query = f"UPDATE assets SET {set_clause} WHERE asset_id = ?"
                values.append(asset_tag)
                cursor.execute(query, values)
                logger.info(f"Updated asset {asset_tag}")
            else:
                # Insert new asset - Map asset_tag to asset_id for the database schema
                asset_data_for_db = asset_data.copy()
                asset_data_for_db['asset_id'] = asset_tag
                if 'asset_tag' in asset_data_for_db and 'asset_tag' != 'asset_id':
                    del asset_data_for_db['asset_tag']
                    
                placeholders = ", ".join(["?"] * len(asset_data_for_db))
                columns = ", ".join(asset_data_for_db.keys())
                query = f"INSERT INTO assets ({columns}) VALUES ({placeholders})"
                cursor.execute(query, list(asset_data_for_db.values()))
                logger.info(f"Inserted new asset {asset_tag}")
            
            conn.commit()
            return True
        except Exception as e:
            logger.error(f"Database error updating asset: {str(e)}")
            return False
        finally:
            conn.close()
    
    def record_scan(self, asset_id, status, tech_name, notes=""):
        """Record a scan event in history"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
                INSERT INTO scan_history (asset_id, status, tech_name, notes)
                VALUES (?, ?, ?, ?)
            ''', (asset_id, status, tech_name, notes))
            
            conn.commit()
            logger.info(f"Recorded scan for {asset_id}: status={status}")
            return True
        except Exception as e:
            logger.error(f"Error recording scan: {str(e)}")
            return False
        finally:
            conn.close()
    
    def get_current_inventory(self, include_deleted=False):
        """Get all assets currently checked in"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        query = '''
            SELECT a.*, h.timestamp as check_in_date
            FROM assets a
            JOIN scan_history h ON a.asset_id = h.asset_id
            WHERE h.id IN (
                SELECT MAX(id) FROM scan_history
                GROUP BY asset_id
            )
            AND h.status = 'in'
        '''
        
        # Add filter for deleted assets if requested
        if not include_deleted:
            query += " AND (a.operational_status IS NULL OR a.operational_status != 'DELETED')"
            
        query += " ORDER BY h.timestamp DESC"
        
        cursor.execute(query)
        
        result = [dict(row) for row in cursor.fetchall()]
        conn.close()
        
        return result
    
    def get_recent_history(self, days=30):
        """Get scan history for the last X days"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT h.*, a.serial_number
            FROM scan_history h
            JOIN assets a ON h.asset_id = a.asset_id
            WHERE h.timestamp >= datetime('now', ?)
            ORDER BY h.timestamp DESC
        ''', (f'-{days} days',))
        
        result = [dict(row) for row in cursor.fetchall()]
        conn.close()
        
        return result
    
    def search_asset_history(self, search_term):
        """Search full history for an asset"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT h.*
            FROM scan_history h
            JOIN assets a ON h.asset_id = a.asset_id
            WHERE a.asset_id LIKE ? OR a.serial_number LIKE ?
            ORDER BY h.timestamp DESC
        ''', (f'%{search_term}%', f'%{search_term}%'))
        
        result = [dict(row) for row in cursor.fetchall()]
        conn.close()
        
        return result
        
    def get_asset_by_id(self, asset_id):
        """Get a specific asset by ID"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM assets WHERE asset_id = ?", (asset_id,))
        
        result = cursor.fetchone()
        if result:
            result = dict(result)
            
        conn.close()
        return result
        
    def get_asset_by_serial(self, serial):
        """Get a specific asset by serial number"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute("SELECT * FROM assets WHERE serial_number = ?", (serial,))
        
        result = cursor.fetchone()
        if result:
            result = dict(result)
            
        conn.close()
        return result
        
    def get_asset_history(self, asset_id, limit=5):
        """Get recent history for a specific asset"""
        conn = sqlite3.connect(self.db_path)
        conn.row_factory = sqlite3.Row
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT * FROM scan_history
            WHERE asset_id = ?
            ORDER BY timestamp DESC
            LIMIT ?
        ''', (asset_id, limit))
        
        result = [dict(row) for row in cursor.fetchall()]
        conn.close()
        
        return result
        
    def delete_asset(self, asset_id):
        """Soft delete an asset (mark as deleted instead of removing)"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            # Update asset to mark as deleted
            cursor.execute("""
                UPDATE assets 
                SET operational_status = 'DELETED', 
                    comments = CASE
                        WHEN comments IS NULL OR comments = '' THEN 'Asset deleted from inventory'
                        ELSE comments || ' | Asset deleted from inventory'
                    END
                WHERE asset_id = ?
            """, (asset_id,))
            
            conn.commit()
            logger.info(f"Soft deleted asset {asset_id}")
            return True
        except Exception as e:
            logger.error(f"Error soft deleting asset: {str(e)}")
            return False
        finally:
            conn.close()
            
    def hard_delete_asset(self, asset_id):
        """Actually delete an asset from the database (for administrator use)"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            # Delete the asset
            cursor.execute("DELETE FROM assets WHERE asset_id = ?", (asset_id,))
            
            # Note: We keep the history records for auditing purposes
            
            conn.commit()
            logger.info(f"Hard deleted asset {asset_id}")
            return True
        except Exception as e:
            logger.error(f"Error hard deleting asset: {str(e)}")
            return False
        finally:
            conn.close()
