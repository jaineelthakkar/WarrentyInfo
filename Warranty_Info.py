# Import the necessary libraries
import pandas as pd
import requests
import time
from bs4 import BeautifulSoup
from selenium import webdriver

# Read the Excel file into a pandas DataFrame
df = pd.read_excel('Book1.xlsx')

# Create an empty list to store the results
results = []

# Extract the serial numbers from the DataFrame
serial_numbers = df['Serial Number'].tolist()

# Iterate over the serial numbers
for i in serial_numbers:
    # Construct the URL for the warranty page
    url = f'https://pcsupport.lenovo.com/us/en/products/laptops-and-netbooks/thinkpad-t-series-laptops/thinkpad-t490-type-20n2-20n3/20n2/20n20030us/{i}/warranty'
    # Create a webdriver instance
    driver = webdriver.Chrome()

    # Navigate to the URL
    driver.get(url)

    # Wait for the page to load and the JavaScript to execute
    time.sleep(5)  # pause for 5 seconds

    # Get the HTML code of the page
    html = driver.page_source

    # Parse the HTML code
    soup = BeautifulSoup(html, 'html.parser')

    # Find the div element with the class "prod-warranty-section"
    div_element = soup.find('div', class_='detail-properties')

    if div_element:
        data = div_element.get_text()
        
        end_date = data[72:82]
    else:
        #print(url)
        data = ''


   
    # Close the browser window
    driver.quit()

  
    # Append the result to the list
    results.append({'Serial Number': i, 'Warranty End Date': end_date})

# Create a DataFrame from the results
df_results = pd.DataFrame(results)

# Write the DataFrame to an Excel file
df_results.to_excel('results.xlsx', index=False)
