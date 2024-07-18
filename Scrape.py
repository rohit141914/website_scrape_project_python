import csv
import time
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# Path to the Edge WebDriver executable
webdriver_path = "msedgedriver.exe"

# Configure Edge WebDriver to run headless
options = Options()
options.add_argument('--headless')
options.add_argument('--disable-gpu')
options.add_argument('--window-size=1920,1080')

# Initialize the WebDriver service
service = Service(webdriver_path)
driver = webdriver.Edge(service=service, options=options)

# Retry configuration
retry_attempts = 3  # Number of attempts to retry loading the page
retry_wait_time = 10  # Wait time in seconds between attempts

try:
    for attempt in range(1, retry_attempts + 1):
        print("Attempting to open the website...")
        driver.get("https://hprera.nic.in/PublicDashboard")

        try:
            # Wait for the page to load completely by looking for a specific element
            WebDriverWait(driver, 60).until(
                EC.presence_of_element_located((By.CSS_SELECTOR, "a[onclick^='tab_project_main_ApplicationPreview']"))
            )
            print("Page loaded successfully.")
            break  # Exit loop if the page loads successfully
        except Exception as e:
            print(f"Attempt {attempt}: Page load timed out. Retrying...")
            print(e)
            if attempt < retry_attempts:
                time.sleep(retry_wait_time)
            else:
                print("Exceeded retry attempts. Please re-run the program.")
                driver.quit()
                exit()

    # Locate project cards on the page
    print("Locating project cards...")
    cards = driver.find_elements(By.CSS_SELECTOR, "a[onclick^='tab_project_main_ApplicationPreview']")
    
    # Process only the first 6 project cards
    print(f"Found {len(cards)} projects. Processing the first 6 projects.")

    data = []

    for index, card in enumerate(cards[:6]):
        print(f"Opening project {index + 1} details...")
        card.click()

        try:
            # Wait for the project details to load
            WebDriverWait(driver, 10).until(
                EC.presence_of_element_located((By.XPATH, "//td[text()='Name']/following-sibling::td"))
            )
            
            # Extract project details
            name = driver.find_element(By.XPATH, "//td[text()='Name']/following-sibling::td").text
            pan_no = driver.find_element(By.XPATH, "//td[text()='PAN No.']/following-sibling::td/span").text
            gstin_no = driver.find_element(By.XPATH, "//td[text()='GSTIN No.']/following-sibling::td/span").text
            address = driver.find_element(By.XPATH, "//td[text()='Permanent Address']/following-sibling::td/span").text

            project_data = {
                "Name": name,
                "PAN No.": pan_no,
                "GSTIN No.": gstin_no,
                "Permanent Address": address
            }

            data.append(project_data)
        except Exception as e:
            print(f"Error extracting data from project {index + 1}: {e}")

        try:
            # Close the project details view
            close_button = driver.find_element(By.XPATH, "//button[text()='Close']")
            if close_button:
                close_button.click()
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "a[onclick^='tab_project_main_ApplicationPreview']"))
                )
            else:
                driver.back()
                WebDriverWait(driver, 5).until(
                    EC.presence_of_element_located((By.CSS_SELECTOR, "a[onclick^='tab_project_main_ApplicationPreview']"))
                )
            print(f"Closed details view for project {index + 1}")
        except Exception as e:
            print(f"Error closing details view for project {index + 1}: {e}")

    # Write the extracted data to a CSV file
    csv_file = 'output.csv'
    with open(csv_file, 'w', newline='') as file:
        writer = csv.DictWriter(file, fieldnames=data[0].keys())
        writer.writeheader()
        writer.writerows(data)

    print(f'CSV file "{csv_file}" has been created successfully.')

    # Print the extracted data
    print("\nExtracted data from the first 6 projects:")
    for project in data:
        print(project)

finally:
    # Close the WebDriver
    print("Closing the browser...")
    driver.quit()
