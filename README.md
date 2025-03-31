# IT Bench Inventory System

A self-contained inventory management system for IT workbench, with ServiceNow CMDB integration.

## Features

- Scan asset tags or enter serial numbers to look up device info
- Automatically retrieves asset details from ServiceNow CMDB
- Track items in/out of inventory
- View current inventory and history
- Search for specific assets
- Export inventory data to CSV

## Installation

1. Ensure you have Python 3.8+ installed
2. Install the required packages:
	"pip install -r requirements.txt"
3. Ensure Microsoft Edge is installed (required for ServiceNow scraping)

## Usage

1. Run the application:
	"python main.py"
2. Scan or enter an asset tag or serial number
3. The application will automatically retrieve the asset details from ServiceNow
4. Choose to check the item IN or OUT of inventory
5. View current inventory in the "Current Inventory" tab
6. View recent activity in the "Recent History" tab

## Deployment to Other Technicians

Option 1: Python Installation
1. Ensure Python 3.8+ is installed on their machine
2. Copy the entire folder to their computer
3. Run "pip install -r requirements.txt"
4. Launch with `python main.py`

Option 2: Executable Deployment
1. Build an executable:
	-"pip install pyinstaller"
	-"pyinstaller --onefile --windowed --name ITBenchInventory --add-data "msedgedriver.exe;." main.py"
2. Copy the generated .exe from the `dist` folder to their computer
3. No additional installation required

## First-Time Use

1. When you first run the application and scan an asset, Edge will open to the ServiceNow login page
2. Complete the Microsoft authentication process
3. Your login session will be saved for future use
4. Subsequent scans should use the saved session unless it expires

Note: If you encounter authentication timeouts, manually log in to ServiceNow in Edge, then try using the application again.

## Troubleshooting

- Ensure Edge browser is installed
- Check that ServiceNow is accessible and the user is logged in
- Logs are stored in `inventory_app.log`