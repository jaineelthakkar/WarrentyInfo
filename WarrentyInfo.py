# Import the necessary libraries
import pandas as pd
import requests
import time
from bs4 import BeautifulSoup
from selenium import webdriver

# Read the Excel file into a pandas DataFrame
df = pd.read_excel('filename.xlsx')

# Extract the serial numbers from the DataFrame
serial_numbers = df['Serial Number'].tolist()

# Iterate over the serial numbers
for serial_number in serial_numbers:
    # Construct the URL for the warranty page
    url = f'https://pcsupport.lenovo.com/us/en/products/laptops-and-netbooks/{serial_number}/warranty'

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

    # Extract the data inside the div element
    data = div_element.get_text()

    # Close the browser window
    driver.quit()
    manu = data[62:82]
    print(manu)
