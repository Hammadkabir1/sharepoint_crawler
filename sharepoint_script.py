import os
import pandas as pd
import re
from urllib.parse import quote
from dotenv import load_dotenv
from office365.runtime.auth.user_credential import UserCredential
from office365.sharepoint.client_context import ClientContext

# ----------------- SHAREPOINT ACCESS FUNCTIONS -----------------

def retrieve_sharepoint_images():
    """Connect to SharePoint and retrieve images from the GPT-Images folder"""
    
    print("Connecting to SharePoint and retrieving images...")
    
    # Check if the file already exists
    excel_file = "SharePoint_GPT_Images.xlsx"
    if os.path.exists(excel_file):
        print(f"Found existing file {excel_file}. Using it instead of retrieving from SharePoint.")
        try:
            images_df = pd.read_excel(excel_file)
            print(f"Loaded {len(images_df)} images from the existing file.")
            return images_df
        except Exception as e:
            print(f"Error reading existing file: {str(e)}")
            print("Will try to retrieve from SharePoint instead.")

    # Load environment variables from the .env file
    load_dotenv()

    # Retrieve the environment variables
    USERNAME = os.getenv('sharepoint_email')
    PASSWORD = os.getenv('sharepoint_password')
    SITE_URL = os.getenv('sharepoint_url_site')
    
    # Remove trailing slash from URL if present
    if SITE_URL.endswith('/'):
        SITE_URL = SITE_URL[:-1]

    try:
        # Get site context
        ctx = ClientContext(SITE_URL).with_credentials(
            UserCredential(USERNAME, PASSWORD)
        )
        
        # Get web property to determine server relative URL base
        web = ctx.web
        ctx.load(web, ["ServerRelativeUrl"])
        ctx.execute_query()
        
        site_relative_url = web.properties["ServerRelativeUrl"]
        print(f"Site Relative URL: {site_relative_url}")
        
        # Construct the proper relative path for GPT-Images folder
        relative_folder_path = f"{site_relative_url}/Branding files/Reference Repository/Reference Repository/GPT-Images"
        print(f"Accessing folder: {relative_folder_path}")

        # Access the folder (GPT-Images)
        folder = ctx.web.get_folder_by_server_relative_url(relative_folder_path)

        # Fetch the files in the folder
        print("Loading files from SharePoint folder...")
        files = folder.files
        ctx.load(files)
        ctx.execute_query()
        print(f"Loaded {len(files)} files from SharePoint")

        # Create lists to store file data
        file_data = []

        # Extract domain and site name for URL construction
        domain_parts = SITE_URL.split('//')
        domain = domain_parts[1].split('/')[0]  # e.g., tmccentral.sharepoint.com
        
        # Get site name from the URL
        site_name = ""
        if "/sites/" in SITE_URL:
            site_name = SITE_URL.split("/sites/")[1].split("/")[0]
        else:
            site_name = domain.split('.')[0]
        
        print(f"Domain: {domain}, Site name: {site_name}")
        
        # Process files in batches to avoid timeouts
        batch_size = 50
        total_files = len(files)
        processed_count = 0
        
        for batch_start in range(0, total_files, batch_size):
            batch_end = min(batch_start + batch_size, total_files)
            print(f"Processing files {batch_start+1} to {batch_end} of {total_files}...")
            
            # Process each file in the batch
            for i in range(batch_start, batch_end):
                file = files[i]
                file_name = file.properties["Name"]
                server_relative_url = file.properties['ServerRelativeUrl']
                
                # Build the proper URL using the Web URL format that works
                # Web URL format: {SITE_URL}/Branding%20files/Reference%20Repository/Reference%20Repository/GPT-Images/{file_name}
                web_url = f"{SITE_URL}/Branding%20files/Reference%20Repository/Reference%20Repository/GPT-Images/{quote(file_name)}"
                
                file_data.append({
                    "File Name": file_name,
                    "File URL": web_url
                })
                
                processed_count += 1
                if processed_count % 10 == 0:
                    print(f"Processed {processed_count} of {total_files} files...")
            
            # Save progress after each batch
            df = pd.DataFrame(file_data)
            df.to_excel(excel_file, index=False)
            print(f"Saved progress: {len(file_data)} files processed")
        
        # Final save to Excel
        print(f"\nCompleted processing {len(file_data)} files.")
        print(f"Image data saved to {excel_file}")
        
        return df

    except Exception as e:
        print(f"An error occurred while accessing SharePoint: {str(e)}")
        print("Please verify your SharePoint URL and credentials")
        return None

# ----------------- REGEX MATCHING FUNCTIONS -----------------

def clean_industry_name(industry):
    """Standardize industry names for better matching"""
    if not isinstance(industry, str):
        return ""
    
    # Convert to lowercase for case-insensitive matching
    industry = industry.lower().strip()
    
    # Handle common variations and abbreviations
    replacements = {
        "oil & gas": "oil&gas",
        "oil and gas": "oil&gas",
        "oil&gas": "oil&gas",
        "oil": "oil&gas",
        "food & beverage": "food&bev",
        "food and beverage": "food&bev",
        "food&bev": "food&bev",
        "food": "food&bev",
        "banking & finance": "bankingfinance",
        "banking and finance": "bankingfinance",
        "banking": "bankingfinance",
        "finance": "bankingfinance",
        "consumer products": "consumerproducts",
        "consumer goods": "consumerproducts",
        "it services": "itservices",
        "it service": "itservices",
        "higher education": "education",
        "telecommunications": "telecom",
        "telecommunication": "telecom",
        "public sector": "government",
        "public sector/govt": "government",
        "public sector/finance": "government",
        "automobile": "auto",
        "automotive": "auto",
        "utilities": "energy",
        "utility": "energy",
        "power transmission": "energy",
        "pharma": "pharmaceutical",
        "pharmaceuticals": "pharmaceutical",
    }
    
    # Apply replacements
    for key, value in replacements.items():
        if industry == key:
            return value
    
    # Remove spaces and special characters for matching
    industry = re.sub(r'[^a-z0-9]', '', industry)
    return industry

def extract_info_from_filename(filename):
    """Extract information from the image filename"""
    # Common pattern in filenames: Product_Industry_CompanyName.jpg
    # Example: SAPS4HANA_Oil&Gas_PSO.jpg or Qlik_SLA_Conglomerate_IBL.jpg
    # or Qlik_Licenses_Retail_Imtiaz.jpg
    
    # Remove file extension
    basename = os.path.splitext(filename)[0]
    
    # Split by underscore
    parts = basename.split('_')
    
    # Default values
    product = ""
    company_name = ""
    industry = ""
    scope = ""
    
    # Extract information based on parts count
    if len(parts) >= 4:  # Format like Qlik_Licenses_Retail_Imtiaz or Qlik_SLA_Conglomerate_IBL
        product = parts[0]
        
        # Special handling for common patterns
        if "license" in parts[1].lower() or "licenses" in parts[1].lower():
            scope = parts[1]
            if len(parts) > 3:  # We have Product_Licenses_Industry_Company
                industry = parts[2]
                company_name = parts[3]
        else:
            # Standard format: Product_Scope_Industry_Company
            scope = parts[1]
            industry = parts[2]
            company_name = parts[3]
            
    elif len(parts) >= 3:  # Format like SAPS4HANA_Oil&Gas_PSO
        product = parts[0]
        industry = parts[1]
        company_name = parts[2]
    elif len(parts) == 2:  # Format like Product_CompanyName
        product = parts[0]
        company_name = parts[1]
    
    # Clean up the extracted industry
    industry = clean_industry_name(industry)
    
    # Handle special case where the company name might be in the last part
    if len(parts) > 1 and not company_name:
        company_name = parts[-1]
    
    return {
        'product': product,
        'scope': scope,
        'industry': industry,
        'company_name': company_name,
        'all_parts': parts  # Include all parts for additional matching
    }

def find_matching_company(extracted_info, sales_df, distinct_industries):
    """Find the matching company in the sales dataframe based on extracted info"""
    company_matches = []
    score_threshold = 0.35  # Lower threshold to catch more matches
    
    # Get the extracted information
    extracted_company = extracted_info['company_name']
    extracted_industry = extracted_info['industry']
    extracted_product = extracted_info['product']
    extracted_scope = extracted_info['scope']
    all_parts = extracted_info['all_parts']
    
    # Create a combined string of all parts for full text matching
    all_parts_text = ' '.join(all_parts).lower()
    
    # Special case handling for retail and common industries
    retail_keywords = ['retail', 'shop', 'store', 'market', 'mart', 'supermarket']
    is_retail_image = any(keyword in all_parts_text for keyword in retail_keywords)
    
    # Search through each row in the sales dataframe
    for index, row in sales_df.iterrows():
        match_score = 0.0
        match_details = []
        
        # Special case for retail matching
        if is_retail_image and pd.notna(row['Industry']) and 'retail' in row['Industry'].lower():
            match_score += 0.2
            match_details.append("Retail industry match")
        
        # ----- COMPANY NAME MATCHING -----
        if pd.notna(row['CompanyName']) and extracted_company:
            company_name = str(row['CompanyName'])
            # Normalize both strings
            company_normalized = re.sub(r'[^a-zA-Z0-9]', '', company_name.lower())
            extracted_normalized = re.sub(r'[^a-zA-Z0-9]', '', extracted_company.lower())
            
            # Check for exact name match - higher priority
            if extracted_normalized.lower() == company_normalized.lower():
                match_score += 0.6
                match_details.append("Exact company name match")
            # Check for company name match
            elif extracted_normalized in company_normalized:
                # Direct match
                match_score += 0.4
                match_details.append("Company name contains abbreviation")
            elif len(extracted_normalized) >= 3 and company_normalized.startswith(extracted_normalized):
                # Starts with the abbreviation
                match_score += 0.3
                match_details.append("Company name starts with abbreviation")
            elif len(extracted_normalized) >= 3 and extracted_normalized in company_normalized:
                # Contains the abbreviation
                match_score += 0.2
                match_details.append("Company name contains partial match")
            
            # Check for acronym match (e.g., IBL for International Brands Limited)
            company_words = company_name.split()
            if len(company_words) > 1:
                acronym = ''.join([word[0].lower() for word in company_words if word[0].isalpha()])
                if acronym == extracted_normalized.lower():
                    match_score += 0.5
                    match_details.append("Matched company acronym")
            
            # Check for abbreviation in the company name
            # E.g., "International Brands (Private) Limited" -> "IBL"
            company_abbr = ''.join([word[0].upper() for word in re.findall(r'\b[a-zA-Z]+\b', company_name)])
            if company_abbr.lower() == extracted_company.lower():
                match_score += 0.5
                match_details.append("Matched company abbreviation")
                
            # Check for "contains word" match - check if the company name contains the extracted company name as a word
            company_words = re.findall(r'\b\w+\b', company_name.lower())
            extracted_words = re.findall(r'\b\w+\b', extracted_company.lower())
            for extracted_word in extracted_words:
                if len(extracted_word) >= 3 and extracted_word in company_words:
                    match_score += 0.35
                    match_details.append(f"Company name contains word '{extracted_word}'")
                    break
                    
            # Check for partial company name match - if company name has multiple words
            company_words = company_name.lower().split()
            for word in company_words:
                if len(word) >= 4 and word in extracted_company.lower():
                    match_score += 0.25
                    match_details.append(f"Company word '{word}' in filename")
                    break
        
        # ----- INDUSTRY MATCHING -----
        if pd.notna(row['Industry']) and extracted_industry:
            sales_industry = clean_industry_name(row['Industry'])
            
            # Check for industry match
            if sales_industry == extracted_industry:
                match_score += 0.3
                match_details.append("Industry exact match")
            elif extracted_industry in sales_industry or sales_industry in extracted_industry:
                match_score += 0.2
                match_details.append("Industry partial match")
        
        # ----- SCOPE/LICENSES MATCHING -----
        if pd.notna(row.get('Description', '')) and "license" in all_parts_text:
            description = str(row['Description']).lower()
            if "license" in description:
                match_score += 0.3
                match_details.append("License match in description") 
        
        # Check Scope separately
        if pd.notna(row.get('Scope', '')) and extracted_scope:
            scope = str(row['Scope']).lower()
            
            if extracted_scope.lower() in scope:
                match_score += 0.2
                match_details.append("Scope match")
        
        # ----- PRODUCT MATCHING -----
        if pd.notna(row.get('Description', '')) and extracted_product:
            description = str(row['Description']).lower()
            
            if extracted_product.lower() in description:
                match_score += 0.2
                match_details.append("Product match in description")
        
        # ----- FULL TEXT MATCHING -----
        # Check if any of the filename parts appear in any of the text fields
        if pd.notna(row.get('Description', '')):
            description = str(row['Description']).lower()
            # Check if any part of the filename is in the description
            for part in all_parts:
                if len(part) >= 3 and part.lower() in description:
                    match_score += 0.15
                    match_details.append(f"Filename part '{part}' in description")
                    break
        
        if pd.notna(row.get('Scope', '')):
            scope = str(row['Scope']).lower()
            # Check if company abbreviation is in the scope
            if extracted_company.lower() in scope:
                match_score += 0.25
                match_details.append("Company abbreviation in scope")
            
            # Check if any part of the filename is in the scope
            for part in all_parts:
                if len(part) >= 3 and part.lower() in scope:
                    match_score += 0.15
                    match_details.append(f"Filename part '{part}' in scope")
                    break
        
        # Add to matches if score is above threshold
        if match_score >= score_threshold:
            match_strength = 'low'
            if match_score >= 0.7:
                match_strength = 'high'
            elif match_score >= 0.5:
                match_strength = 'medium'
            
            company_matches.append({
                'index': index,
                'company': row['CompanyName'] if pd.notna(row['CompanyName']) else "",
                'industry': row['Industry'] if pd.notna(row['Industry']) else "",
                'match_strength': match_strength,
                'match_score': match_score,
                'match_details': match_details
            })
    
    # Sort matches by score (highest first)
    company_matches.sort(key=lambda x: x['match_score'], reverse=True)
    
    return company_matches

def match_images_to_companies(images_df, sales_file="Sales_Compiled_Sheet 1.xlsx"):
    """Match image files to companies in the sales sheet"""
    
    print("\nStarting image matching process...")
    
    # Check if the sales file exists
    if not os.path.exists(sales_file):
        print(f"Error: Sales file '{sales_file}' not found.")
        return 0
    
    # Load the sales Excel file
    print(f"Loading sales sheet from {sales_file}...")
    try:
        sales_df = pd.read_excel(sales_file)
    except Exception as e:
        print(f"Error reading sales file: {str(e)}")
        return 0
    
    # Get distinct industries from the master excel
    distinct_industries = sales_df['Industry'].dropna().unique()
    print(f"Found {len(distinct_industries)} distinct industries in the sales sheet.")
    
    # Identify the column for reference images URL
    reference_column = "Reference Images' URL"
    if reference_column not in sales_df.columns:
        print(f"Warning: Column '{reference_column}' not found in sales sheet. It will be created.")
        sales_df[reference_column] = ""
    
    # Ensure the reference column is string type to avoid warnings
    sales_df[reference_column] = sales_df[reference_column].astype(str)
    
    # Process each image file
    print("Processing image files...")
    matches_found = 0
    matches_details = []
    
    for index, row in images_df.iterrows():
        # Extract information from filename
        filename = row['File Name']
        file_url = row['File URL']
        
        if index % 50 == 0 or index == 0:
            print(f"\nProcessing file {index+1}/{len(images_df)}: {filename}")
            
        extracted_info = extract_info_from_filename(filename)
        
        # Find matching companies
        matches = find_matching_company(extracted_info, sales_df, distinct_industries)
        
        if matches:
            # Update the URL for the best match
            best_match = matches[0]
            sales_df.at[best_match['index'], reference_column] = file_url
            matches_found += 1
            
            # Add details for logging
            match_detail = {
                'file_name': filename,
                'matched_company': best_match['company'],
                'industry': best_match['industry'],
                'match_score': best_match['match_score'],
                'match_details': best_match['match_details']
            }
            matches_details.append(match_detail)
            
            if index % 50 == 0 or index == 0:
                print(f"  Found match: {best_match['company']} - {best_match['industry']}")
                print(f"  Match score: {best_match['match_score']:.2f}, Details: {', '.join(best_match['match_details'])}")
    
    # Save the updated sales sheet
    output_file = "Sales_Compiled_Sheet_Updated.xlsx"
    print(f"\nSaving updated sales sheet to {output_file}...")
    try:
        sales_df.to_excel(output_file, index=False)
        print(f"Updated sales sheet saved to {output_file}")
    except Exception as e:
        print(f"Error saving updated sales sheet: {str(e)}")
        print("Try saving the file manually or check file permissions.")
    
    # Save matching details to a log file for review
    details_df = pd.DataFrame(matches_details)
    details_df.to_excel("Matching_Details_Log.xlsx", index=False)
    print(f"Match details saved to Matching_Details_Log.xlsx for review")
    
    print(f"Process completed. Found matches for {matches_found} out of {len(images_df)} images.")
    
    return matches_found

# ----------------- MAIN FUNCTION -----------------

def main():
    print("Starting SharePoint Image Processing")
    print("=" * 50)
    
    # Step 1: Retrieve images from SharePoint
    images_df = retrieve_sharepoint_images()
    
    if images_df is None or len(images_df) == 0:
        print("No images were retrieved from SharePoint. Exiting.")
        return
    
    # Step 2: Match images to companies in the sales sheet
    match_images_to_companies(images_df)
    
    print("\nProcess completed successfully!")

if __name__ == "__main__":
    main()
