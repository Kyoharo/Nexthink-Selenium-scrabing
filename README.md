# NxThink - Automate NXPortal Data Extraction (Python)

###Description:

This Python script automates the process of extracting data from the NXPortal system (presumably used by Egypt Post). It utilizes Selenium for web automation and interacts with the NXPortal interface to gather information about device issues. The script can navigate through different menus and services, extract relevant data, and save it to an Excel spreadsheet.

###Key Features:

- Automates NXPortal data extraction
- Handles date navigation within the interface
- Extracts device issues for various services
- Saves data to an Excel spreadsheet

### Getting Started:

- Install required Python libraries `pip install -r requirements.txt`.
- Set environment variables for `username`, `password`, and file paths using a `.env` file (refer to Docker instructions for details).
- Run the script: `python new_nxthink.py`

### Getting Started with Docker:
- Build docker image using: `docker build -t nxthink_app .`
- Run the image using: `docker run -e NAME=your_username -e PASSWD=your_password nxthink_app`

