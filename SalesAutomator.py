import os
import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.chrome.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from dotenv import load_dotenv
from datetime import datetime

# Generate a date range (YYYY-MM-DD)
dates = pd.date_range(start='2024-11-06', end='2024-11-11', freq='D')

# Country's ID's
ids = {
    'Argentina'     : 1
    ,'Brasil'       : 2
    ,'Costa Rica'   : 3
    ,'Ecuador'      : 4
    ,'Perú'         : 5
    ,'Uruguay'      : 6
}

# Load environment variables from the .env file
load_dotenv('security.env')
BP_USER = os.getenv('BP_USER')
BP_PASSWORD = os.getenv('BP_PASSWORD')
BP_URL = os.getenv('BP_URL')
BP_URL_FILTER = os.getenv('BP_URL_FILTER')

# Configure Chrome browser options
chrome_options = Options()
chrome_options.add_argument("user-data-dir=/path/to/your/chrome/profile")

# Open Chrome browser
driver = webdriver.Chrome(options=chrome_options)
driver.get(BP_URL)

# Waiting time for the username and passwor<Z+d fields to appear
wait = WebDriverWait(driver, 2)

# To store results
data = []

try:
    # Enter the username
    username_field = wait.until(EC.visibility_of_element_located((By.ID, 'email_address_login')))
    username_field.send_keys(BP_USER)

    # Enter the password
    password_field = wait.until(EC.visibility_of_element_located((By.ID, 'password')))
    password_field.send_keys(BP_PASSWORD)

    # Click the login button
    login_button = wait.until(EC.element_to_be_clickable((By.ID, 'submit-admin')))
    login_button.click()

    # Login reference point
    wait.until(EC.presence_of_element_located((By.ID, 'utility')))
    print(">>>>>Successful login to Brightpearl")
    print(f'>>>> Bucle start at {datetime.now()}')

    # Loop through dates and IDs
    for date in dates:
        date_str = date.strftime('%d%%2F%m%%2F%Y')
        for country, id in ids.items():
            url = BP_URL_FILTER.format(date_from=date_str, date_to=date_str, ids=id)
            driver.get(url)
            #print(f"Loaded URL for {country} on {date}")
            wait

            # Extract the full text of the element
            try:
                pagination_info = wait.until(EC.visibility_of_element_located((By.ID, 'sales_by_channel-total_order_sales-foot')))
                full_text = pagination_info.text.strip()
                data.append({'Date': date, 'Country' : country, 'Sales': full_text})
                wait
            except Exception as e:
                print(f"Error extracting sales data for {country} on {date}: {str(e)}")
                data.append({'Date': date, 'Sales': 'Error'})
                wait

except Exception as e:
    print(f">>>>>Error: {str(e)}")

finally:
    print(f'>>>>>Process finished at {datetime.now()}')
    driver.quit()

# Create DataFrame
df = pd.DataFrame(data).sort_values(by='Date', ascending=True)
# Ensure 'Date' is in datetime format
df['Date'] = pd.to_datetime(df['Date'])
# Format the min and max dates from the dataframe
min_date_str = df['Date'].min().strftime("%Y_%m_%d")
max_date_str = df['Date'].max().strftime("%Y_%m_%d")
# Save the CSV with formatted dates
df.to_csv(f'{datetime.now().strftime("%Y_%m_%d")}_Sales_{min_date_str}_{max_date_str}.csv', index=False)

