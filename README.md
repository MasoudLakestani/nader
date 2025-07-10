Gold Transaction Manager - Windows Version
A simple Windows application for managing gold transactions with Persian/English interface.

🚀 Quick Start
Option 1: Run Python Script Directly
Install Python (if not already installed):
Download from https://www.python.org/downloads/
During installation, check "Add Python to PATH"
Run Setup:
Double-click setup.bat
This will install required packages
Run the Application:
Double-click run.bat or
Run: python gold_app.py
Option 2: Build Standalone EXE
Complete Option 1 first
Build EXE:
Double-click build.bat
Wait for build to complete
Find GoldTransactionManager.exe in the folder
Run EXE:
Double-click GoldTransactionManager.exe
No Python installation needed on other computers!
📁 Files Included
gold_app.py - Main application code
setup.bat - Install Python packages
build.bat - Build standalone EXE
run.bat - Run the application
requirements.txt - Python dependencies
README.md - This file
🔧 Features
✅ Buy Gold Recording - Record gold purchases with weight, karat, and price
✅ Sell Gold Recording - Record gold sales
✅ Inventory Management - View current gold inventory
✅ Transaction History - View all past transactions
✅ Excel Database - All data stored in Excel file
✅ Persian/English Interface - Supports both languages
✅ Auto Excel Creation - Creates Excel file automatically if missing
📊 How it Works
First Run: Application creates transactions.xlsx file
Recording: All transactions are saved to Excel file
Calculations: Automatically converts between grams and methqal
Inventory: Calculates current inventory (purchases - sales)
🛠️ Troubleshooting
Python not found
Install Python from https://www.python.org/downloads/
Make sure to check "Add Python to PATH" during installation
Package installation fails
Run Command Prompt as Administrator
Run: pip install openpyxl pyinstaller
EXE build fails
Make sure all packages are installed
Check that gold_app.py is in the same folder as build.bat
Excel file issues
Make sure you have write permissions in the folder
Close Excel if the file is open
The application will recreate the file if deleted
📋 System Requirements
OS: Windows 7/8/10/11
Python: 3.7+ (for running Python version)
Memory: 50MB RAM
Storage: 10MB free space
🔄 Conversion Rate
1 Methqal = 4.3317 Grams (used in calculations)
📞 Support
If you encounter any issues:

Check the troubleshooting section above
Make sure Python is properly installed
Verify all files are in the same folder
Try running as Administrator if needed
🎯 Usage Tips
Backup: Regularly backup your transactions.xlsx file
Precision: Enter weights with decimal precision (e.g., 12.50)
Karat: Use standard karat values (750, 900, etc.)
Notes: Add notes for better record keeping
