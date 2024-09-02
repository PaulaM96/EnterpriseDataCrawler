# EnterpriseDataCrawler

## Description

A tool for automating the extraction of company data, including business names, contacts, and registration details, from public sources using Selenium and Excel integration.

## Features

- Automatic extraction of company data such as trade names, contacts, and registration numbers.
- Saves extracted data to an Excel file for easy access and management.
- User-friendly interface for starting and canceling searches.
- Supports multiple search terms and collects data from various pages.
- Built with Python, SeleniumBase, Tkinter, and OpenPyXL.

## Installation

### Prerequisites

- Python 3.8 or above
- Google Chrome
- ChromeDriver (compatible version with Chrome installed)
- Required Python packages

### Installation Steps

1. **Clone the repository and install required python packages:**
   ```bash
   git clone https://github.com/your-username/EnterpriseDataCrawler.git
   cd EnterpriseDataCrawler
   pip install -r requirements.txt´´´

4. **Set up the ChromeDriver:**

Download the correct version of ChromeDriver.
Place the driver in the drivers folder inside the project directory.

### How to Use
**Run the application:**
**Using the Interface:**

- Enter a search term related to the company or business you want to search for.
- Click "Search" to start the data extraction process.
- The progress will be displayed through a progress bar and a URL count.
- Click "Cancel" at any time to stop the search.
- Once the search is complete, click "Open Folder" to access the saved Excel file.
  
**Saved Data:** 
Data is saved in the folder C:\DADOS_CNPJ with a filename format: cnpjs_ativos_{search_term}.xlsx.

### Troubleshooting: 
Ensure ChromeDriver is correctly placed and matches the installed version of Google Chrome.
If the search is too slow, check your internet connection and the response time of the data source.
### Contributing: 
Feel free to fork this repository, open issues, and submit pull requests. Contributions are welcome!
