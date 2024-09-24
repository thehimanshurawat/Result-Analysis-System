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
            result_soup = BeautifulSoup(driver.page_source, 'html.parser')
            table = result_soup.find('table', {'border': '1'})

            if not table:
                print("No result table found on the results page. Please re-enter the CAPTCHA.")
                driver.close()  # Close the current tab
                driver.switch_to.window(driver.window_handles[0])
                continue  # Prompt for CAPTCHA again

            # Parse the result table
            rows = table.find_all('tr')[1:]  # Skipping the header row
            result_data = []
            for row in rows:
                # Extract course code and title
                course_info = row.find_all('td')[1].find_all('td')
                course_code = course_info[0].text.strip()
                course_title = course_info[1].text.strip()
                
                # Extract max marks
                max_marks_info = row.find_all('td')[3].find_all('tr')
                max_marks = {}
                if max_marks_info:
                    for mark_row in max_marks_info:
                        cells = mark_row.find_all('td')
                        if len(cells) == 2:  # Ensure there are two cells for each mark entry
                            max_marks[cells[0].text.strip()] = cells[1].text.strip()
                
                # Extract marks/grades
                grades_info = row.find_all('td')[4].find_all('td')
                marks = [grade.text.strip() for grade in grades_info if grade.text.strip() != '']
                
                # Append the data to result_data as separate columns
                result_data.append({
                    'Course Code': course_code,
                    'Course Title': course_title,
                    'Max Marks Theory': max_marks.get('Theory', ''),
                    'Max Marks Sessional': max_marks.get('Sessional', ''),
                    'Max Marks Practical': max_marks.get('Practical', ''),
                    'Marks/Grade': marks[-1] if marks else ''  # Take the last grade if available
                })
                
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
