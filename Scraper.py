# ======================== THIS PROJECT IS DONE BY SHEHAB HABILA, BIOSTATISTICIAN AND R/PYTHON PROGRAMMER ========================
# ================================================================================================================================
# 
# 1. Install python
# 2. Install the dependencies "selenium", "pandas", "bs4", "os", "time"
# 3. Install the webdriver to enable automatic control of the browser (in my case, using firefox, it's geckodriver)
# 4. Configure the path of the webdriver
# 5. Configure the URL of the results page, as well as the needed parameters you need to grap for each student
# 6. Prepare a file with the IDs of the students, and configure its name withing the code below
# 7. Run the code




#

from selenium import webdriver
from selenium.webdriver.firefox.service import Service
from selenium.webdriver.firefox.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
import pandas as pd
from bs4 import BeautifulSoup
import time
import os



# Path for Geckodriver
geckodriver_path = "/usr/local/bin/geckodriver" # If you're using firefox on linux

# Configuring Selenium to use Firefox
firefox_options = Options()
# firefox_options.add_argument("--headless")  # Uncomment to run in headless mode

service = Service(geckodriver_path)
driver = webdriver.Firefox(service=service, options=firefox_options)

# Load student IDs from Excel
df = pd.read_excel("Students Data.xlsx")  # Ensure the file exists
ids = df["ID"].astype(str).tolist()

# Output file
output_file = "student_results.xlsx"

# Check if the output file exists, if not, create it with headers
if not os.path.exists(output_file):
    pd.DataFrame(columns=["ID", "EGU", "GIT", "Communication Skills", "CNS", "Concepts 1", "Concepts 2" , "Professionalism", "Filler", "Year Total"]).to_excel(output_file, index=False)

# URL of the results page
url = "http://www.med.alexu.edu.eg/results/index.php/2025/07/20/2ndresults2/"

for student_id in ids:
    try:
        # Load the page and wait for it to load
        driver.get(url)
        WebDriverWait(driver, 10).until(EC.presence_of_element_located((By.TAG_NAME, "iframe")))

        # Locate and switch to the iframe
        iframe = driver.find_element(By.TAG_NAME, "iframe")
        driver.switch_to.frame(iframe)

        # Locate the input field and enter the student ID
        input_field = WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.ID, "student_id"))
        )
        input_field.clear()
        input_field.send_keys(round(float(student_id)))

        # Submit the form (assuming it's a button)
        submit_button = WebDriverWait(driver, 10).until(
            EC.element_to_be_clickable((By.CSS_SELECTOR, "button[type='submit']"))
        )
        submit_button.click()

        # Wait for results table to appear
        WebDriverWait(driver, 10).until(
            EC.presence_of_element_located((By.TAG_NAME, "table"))
        )
        time.sleep(2)  # Extra delay for safety

        # Get page source and parse with BeautifulSoup
        soup = BeautifulSoup(driver.page_source, "html.parser")

        # Find all "Total" rows
        total_rows = [row for row in soup.find_all("tr") if "total" in row.text.lower()]

        # Extract values for each category
        total_values = ["Not Found"] * 9  # Default values in case some rows are missing

        for i, name in enumerate(["EGU", "GIT", "Communication Skills", "CNS", "Concepts 1", "Concepts 2" , "Professionalism", "Filler", "Year Total"]):
            if len(total_rows) > i:
                total_values[i] = total_rows[i].find_all("td")[-1].get_text(strip=True)

        print(f"✅ Student ID: {student_id} | EGU: {total_values[0]}, GIT: {total_values[1]}, Communication: {total_values[2]}, CNS: {total_values[3]}, Concept I: {total_values[4]}, Concept II: {total_values[5]}, Professionalism: {total_values[6]}, Year Total: {total_values[8]}")

        # Save to Excel immediately
        new_entry = pd.DataFrame([{
            "ID": student_id, "EGU": total_values[0], "GIT": total_values[1], 
            "Communication": total_values[2], "CNS": total_values[3],
            "Concepts 1": total_values[4], "Concepts 2": total_values[5],
            "Professionalism": total_values[6], "Filler": total_values[7],
            "Year Total": total_values[8]
        }])
        with pd.ExcelWriter(output_file, mode="a", if_sheet_exists="overlay", engine="openpyxl") as writer:
            new_entry.to_excel(writer, index=False, header=False, startrow=writer.sheets["Sheet1"].max_row)

        # Switch back to the main page before the next iteration
        driver.switch_to.default_content()

    except Exception as e:
        print(f"❌ Error with ID {student_id}: {e}")
        new_entry = pd.DataFrame([{
            "ID": student_id, "EGU": "Error", "GIT": "Error", 
            "Communication": "Error", "CNS": "Error",
            "Concepts 1": "Error", "Concepts 2": "Error",
            "Professionalism": "Error", "Filler": "Error",
            "Year Total": "Error"
        }])
        with pd.ExcelWriter(output_file, mode="a", if_sheet_exists="overlay", engine="openpyxl") as writer:
            new_entry.to_excel(writer, index=False, header=False, startrow=writer.sheets["Sheet1"].max_row)

# Close the browser
driver.quit()
print(f"📁 Results updated in {output_file}")
