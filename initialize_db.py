from models.database import InventoryDatabase

# Initialize the database with a config file
db = InventoryDatabase('db_config.json')

# The initialization will create the tables
print("Database initialized successfully")