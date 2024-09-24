from selenium import webdriver
from selenium.webdriver.common.by import By
import time
import pandas as pd
from bs4 import BeautifulSoup

# URLs of the university website
result_submission_url = "https://jcboseustymca.co.in/Forms/Student/ResultStudents.aspx"
result_display_url = "https://jcboseustymca.co.in/Forms/Student/PrintReportCardNew.aspx"

def get_result(roll_no, semester):
    # Set up the browser
    driver = webdriver.Chrome()  # Ensure the path to the WebDriver is set if necessary
    driver.get(result_submission_url)

    # Fill in the Roll Number
    roll_no_input = driver.find_element(By.ID, "txtRollNo")
    roll_no_input.clear()
    roll_no_input.send_keys(roll_no)

    # Select Semester
    semester_select = driver.find_element(By.ID, "ddlSem")
    semester_select.send_keys(semester)

    while True:  # Loop until a valid result is retrieved
        # Display CAPTCHA image to the user
        print("Please enter the CAPTCHA displayed on the website.")
        
        # Wait for user to input CAPTCHA
        captcha_input = input("Enter CAPTCHA: ")
        
        # Fill in the CAPTCHA
        captcha_field = driver.find_element(By.ID, "txtCaptcha")
        captcha_field.clear()
        captcha_field.send_keys(captcha_input)

        # Click the View Result button
        submit_button = driver.find_element(By.ID, "btnResult")
        submit_button.click()

        time.sleep(0.1)  # Wait for the results page to load

        # Check if a new tab is opened
        if len(driver.window_handles) > 1:
            # Switch to the new tab
            driver.switch_to.window(driver.window_handles[1])

            # Use BeautifulSoup to parse the HTML and extract data
            #changes
            result_soup = BeautifulSoup(driver.page_source, 'html.parser')
            table = result_soup.find('table', {'border': '1'})

            if not table:
                print("No result table found on the results page. Please re-enter the CAPTCHA.")
                driver.close()  # Close the current tab
                driver.switch_to.window(driver.window_handles[0])  # Switch back to the original tab
                continue  # Prompt for CAPTCHA again

            # Parse the result table
            #changes
            rows = table.find_all('tr')[1:]  # Skipping the header row
            result_data = []
            for row in rows:
                cols = [col.text.strip() for col in row.find_all('td')]
                result_data.append(cols)

            driver.quit()  # Close the browser after scraping
            return result_data  # Return results if successful
        else:
            print("Failed to open results page. Please re-enter the CAPTCHA.")
            continue  # Prompt for CAPTCHA again

def main():
    year = input("Enter Year (e.g., 2024): ")
    semester = input("Enter Semester (e.g., Sixth Semester): ")

    results = []

    # Assuming roll numbers are known in advance
    roll_numbers = ['21011004001', '21011004002', '21011004003']  # Add actual roll numbers

    for roll_no in roll_numbers:
        result = get_result(roll_no, semester)
        if result:
            results.append({'Roll No': roll_no, 'Result': result})

    # Save results to Excel
    if results:
        df = pd.DataFrame(results)
        file_name = f'results_{year}_{semester}.xlsx'
        df.to_excel(file_name, index=False)
        print(f"Results saved to {file_name}")
    else:
        print("No results to save.")

if __name__ == "__main__":
    main()
