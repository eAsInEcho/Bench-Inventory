import psycopg2
import logging

# Set up logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
logger = logging.getLogger(__name__)

def create_database():
    try:
        # Connect to the default 'postgres' database first
        logger.info("Connecting to PostgreSQL server...")
        conn = psycopg2.connect(
            dbname="postgres",  # Connect to default postgres database
            user="inventory_user",
            password="Globalfoundries2025!",
            host="10.250.2.203",
            port="5432"
        )
        conn.autocommit = True
        cursor = conn.cursor()
        
        # Check if the inventory database already exists
        cursor.execute("SELECT 1 FROM pg_database WHERE datname = 'inventory'")
        exists = cursor.fetchone()
        
        if not exists:
            logger.info("Creating 'inventory' database...")
            cursor.execute("CREATE DATABASE inventory")
            logger.info("Database 'inventory' created successfully!")
        else:
            logger.info("Database 'inventory' already exists.")
        
        cursor.close()
        conn.close()
        logger.info("Connection closed.")
        return True
    except Exception as e:
        logger.error(f"Error creating database: {e}")
        return False

if __name__ == "__main__":
    success = create_database()
    if success:
        print("\nDatabase created successfully! ✅\n")
    else:
        print("\nFailed to create database! ❌\n")