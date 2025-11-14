AutoCrosschex is a utility program that automates the process of managing attendance data from CrossChex devices. 
The program connects to devices, downloads new attendance records, stores them in a configured database, 
and generates monthly Excel reports for each department.

Key Features:
- Automatic retrieval of new attendance records from CrossChex devices.
- Insertion of records into a database.
- Generation of department-specific Excel reports for each month.
- Fully configurable via files in the 'cfg' folder: paths, devices, and database connection settings.

Getting Started:
1. Configure all `.cfg` files in the `cfg` folder:
   - `paths.cfg` – paths for Excel reports per department.
   - `devices.cfg` – IP addresses of CrossChex devices (port 5010, device ID used only for database records).
   - `db.cfg` – database connection parameters (`server`, `database`, `uid`, `password`).
2. Run `AutoCrosschex.exe`. The program will download new records, update the database, and generate an Excel report for the current month.

AutoCrosschex simplifies attendance management and reporting by automating all repetitive tasks.
