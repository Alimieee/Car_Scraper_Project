# CAR SCRAPER PROJECT - SETUP INSTRUCTIONS

This project contains web scrapers for 4 sites: Mudah, Carsome, Carro, and Carlist.

## IF YOU ARE USING A MAC
Since the project was built on a Mac, the required libraries (including the virtual environment) are already included inside the `mudah_site/venv` folder. 
You can run the scrapers directly from your Mac terminal without installing anything new:

1. Open Terminal.
2. Navigate to the project folder, for example:
   cd path/to/Car_Scraper_Project
3. To run Carro (for example), type:
   cd carro_site
   source ../mudah_site/venv/bin/activate
   python3 carro_deep_scraper.py


## IF YOU ARE USING WINDOWS
Virtual environments cannot be transferred from Mac to Windows. You will need to install the required libraries yourself.

1. Ensure Python is installed on your Windows machine (download from python.org).
2. Open Command Prompt or PowerShell and navigate to the `Car_Scraper_Project` folder.
3. Create a new virtual environment:
   python -m venv venv
4. Activate the virtual environment:
   venv\Scripts\activate
5. Install the required libraries:
   pip install -r requirements.txt
6. Install Playwright browser binaries (Required for Carlist):
   playwright install
7. You can now run the scrapers (ensure your venv is activated). For example:
   cd carro_site
   python carro_deep_scraper.py
