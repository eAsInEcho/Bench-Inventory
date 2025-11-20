# IT Bench Inventory Manager

## Overview
IT Bench Inventory Manager is a desktop application designed to track and manage IT assets, allowing easy scanning, check-in/out, and detailed record-keeping of hardware across multiple technician workstations using a centralized database.

## Features
- Asset scanning and tracking
- Check-in/out functionality
- Detailed asset history
- Integration with ServiceNow CMDB
- Export and reporting capabilities
- Multi-technician collaboration with centralized PostgreSQL database
- Automatic failover to replica databases (if configured)

## Installation

### Prerequisites
- Windows 10 or later
- Chrome or Edge browser (for ServiceNow integration)
- Network connectivity to the PostgreSQL database server

### Installation Steps
1. Download the latest release 
2. Unzip the file (if compressed)
3. Ensure database configuration is present in `db_config.json`
4. Double-click `ITBenchInventory.exe` to run

### Database Setup (Administrator)
The application requires a PostgreSQL database server. For administrators setting up the server:

1. Install PostgreSQL 13.2 or later on your server
2. Create a database named "inventory"
3. Create a user with appropriate privileges
4. Configure pg_hba.conf to allow connections from client machines
5. Ensure port 5432 is accessible on the network

### Configuration
The application uses a `db_config.json` file for database connection:

```json
{
    "primary": {
      "dbname": "inventory",
      "user": "your_username",
      "password": "your_password",
      "host": "database_server_ip",
      "port": "5432"
    },
    "replicas": []
}
```

For testing connectivity, run:
```
python test_db_connection.py
```

## Using the Application

### Scanning Assets
- Use the "Scan" tab to input asset tags or serial numbers
- The app supports both manual entry and barcode scanning
- All asset data is stored in the centralized database

### ServiceNow Integration
- Create a bookmark in your browser to extract asset data from ServiceNow
- Follow the in-app instructions for bookmark creation

### Multi-Technician Workflow
- All technicians work with the same database in real-time
- Changes made by one technician are immediately visible to others
- Database connections are pooled for optimal performance

## Packaging (For Developers)
To package the application:

1. Install required dependencies: `pip install -r requirements.txt`
2. Update `db_config.json` with production database details
3. Use PyInstaller to create an executable:
   ```
   pyinstaller --onefile --windowed --add-data "db_config.json;." main.py
   ```
4. Distribute the resulting executable and configuration files

## Troubleshooting
- Ensure network connectivity to the database server
- Verify the database credentials in `db_config.json`
- Check the `inventory_app.log` file for detailed error information
- Ensure PostgreSQL client libraries are properly installed

## License
This software is proprietary and privately distributed. 

### Distribution Terms
- This software may only be distributed by the original author
- Unauthorized copying, distribution, or modification is strictly prohibited
- Use is permitted only with explicit permission from the software's creator

## Version History
- v1.2.0 - Add bulk checkin/out feature
- v1.1.0 - Fix local DB sync with Central, fix/finalize UI
- v1.0.6 - Add auditing feature
- v1.0.5 - Add flag feature for sought assets. Bug Fixes
- v1.0.4 - Simplify time stamps. Clean Check In look on status. Remove edit for notes and make note font smaller.
- v1.0.3 - Add Scan, Your Bench, All Benches, Out, and Recent History tabs.
		-Fixed Site bug. Assets only change site when checked in or out.
- v1.0.2 - Site change depending on where checked in or if checked out
- v1.0.1 - Fallback to local db if no connection to main DB. Indicator of connection status added.
- v1.0.0 - PostgreSQL Database Implementation
- v0.9.0 - Initial Release