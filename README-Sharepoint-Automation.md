# SharePoint Image Matching Solution

## Overview
This solution automatically retrieves images from a SharePoint folder and matches them to company records in a sales spreadsheet based on the image filename. The matched image URLs are then added to the sales spreadsheet, allowing for easy reference to relevant company images.

## Key Features
- Connects to SharePoint using credentials from a `.env` file
- Retrieves images from a specified SharePoint folder
- Performs intelligent matching of images to companies using multiple criteria:
  - Company name matching (exact, partial, acronym, abbreviation)
  - Industry matching
  - Product and scope matching
  - Keyword matching in descriptions
- Creates a detailed matching log for review
- Outputs an updated sales spreadsheet with image URLs

## How It Works

### 1. SharePoint Connection
The script connects to SharePoint using the credentials and URL provided in the `.env` file. It retrieves images from the "GPT-Images" folder in the Reference Repository.

### 2. Image Information Extraction
For each image, the script extracts information from the filename:
- Product (e.g., "Qlik", "SAPS4HANA")
- Industry (e.g., "Retail", "Oil&Gas")
- Company Name (e.g., "Imtiaz", "PSO")
- Scope (e.g., "Licenses", "SLA")

Examples of supported filename formats:
- `Product_Industry_CompanyName.jpg` (e.g., "SAPS4HANA_Oil&Gas_PSO.jpg")
- `Product_Scope_Industry_CompanyName.jpg` (e.g., "Qlik_SLA_Conglomerate_IBL.jpg")
- `Product_Licenses_Industry_CompanyName.jpg` (e.g., "Qlik_Licenses_Retail_Imtiaz.jpg")

### 3. Matching Algorithm
The script uses a sophisticated matching algorithm that considers multiple factors:
- **Company Name Matching**: Checks for exact matches, abbreviations, acronyms, and partial matches
- **Industry Matching**: Standardizes industry names and checks for matches
- **Product & Scope Matching**: Looks for product names and scope in the company descriptions
- **Special Case Handling**: Special handling for retail and other common industries
- **Keyword Matching**: Checks if keywords from the filename appear in company descriptions

Each match is assigned a score based on these factors, and the best match is selected.

### 4. Output Generation
The script produces:
1. `SharePoint_GPT_Images.xlsx` - A list of all images and their URLs from SharePoint
2. `Sales_Compiled_Sheet_Updated.xlsx` - The updated sales spreadsheet with image URLs
3. `Matching_Details_Log.xlsx` - A detailed log of all matches for verification

## How to Use

1. Ensure you have the required Python packages installed:
   ```
   pip install pandas office365 python-dotenv openpyxl
   ```

2. Create a `.env` file with your SharePoint credentials:
   ```
   sharepoint_email=your.email@domain.com
   sharepoint_password=your_password
   sharepoint_url_site=https://your-site.sharepoint.com/sites/your-site
   ```

3. Place your sales data in a file named `Sales_Compiled_Sheet 1.xlsx` with columns:
   - CompanyName
   - Industry
   - Description (optional)
   - Scope (optional)

4. Run the script:
   ```
   python sharepoint_script.py
   ```

## Requirements
- Python 3.6+
- Required packages:
  - pandas
  - office365
  - python-dotenv
  - openpyxl

## Notes
- The script handles a wide variety of filename formats and company naming conventions
- For best results, use consistent naming conventions for image files
- The matching threshold is set to capture a wide range of potential matches
- The detailed log file can be used to verify and troubleshoot matches 