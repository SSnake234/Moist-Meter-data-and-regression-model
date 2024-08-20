from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select
from bs4 import BeautifulSoup
import pandas as pd
import time

# Initialize Selenium WebDriver (Ensure you have the correct driver for your browser)
driver = webdriver.Chrome()  # Or use Firefox, Edge, etc.

# URL to crawl data from
url = 'https://moistmeter.org/#'  # Replace with your target URL

# Open the URL
driver.get(url)

# Allow the page to load fully
time.sleep(3)

# Find the dropdown button by its ID and click it to reveal the options
dropdown_button = driver.find_element(By.ID, 'tableFilter')
dropdown_button.click()

# Allow the dropdown to open
time.sleep(2)

# Now, find the dropdown menu that appeared (assuming it has a known class or parent)
dropdown_menu = driver.find_element(By.CLASS_NAME, 'dropdown-menu')  # Adjust if necessary

# Find all the category options within the dropdown
options = dropdown_menu.find_elements(By.TAG_NAME, 'a')  # Assuming options are links (`<a>` tags)

# Initialize lists to store data
Id = []
dates = [] 
categories = []
titles = []
ratings = []

# Iterate over each option in the dropdown
for i in range(1, len(options)):
    category = options[i].text
    # Click the option to select the category
    options[i].click()
    
    # Allow the page to load data for the selected category
    time.sleep(3)
    
    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(driver.page_source, 'html.parser')
    
    # Find the table body and extract the data
    body = soup.find('tbody')
    if body:
        all_data = body.find_all('tr')
        for data in all_data:
            cols = data.find_all('td')
            if len(cols) >= 4:
                Id.append(cols[0].text.strip())
                dates.append(cols[1].text.strip())
                titles.append(cols[2].text.strip())
                ratings.append(cols[3].text.strip())
                categories.append(category)
        print(f'Done processing {category} category')
    else:
        print(f"No data found for category: {category}")

    # Re-click the dropdown to select the next option
    dropdown_button.click()
    time.sleep(2)

# Close the Selenium driver
driver.quit()

# Create a DataFrame with the data
df = pd.DataFrame({
    'ID': Id,
    'Date': dates,
    'Category': categories,
    'Title': titles,
    'Rating': ratings
})

# Write the data to an Excel file
df.to_excel('moistmeter_by_category.xlsx', index=False)

print("Data has been successfully written to 'moistmeter_by_category.xlsx'")
