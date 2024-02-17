import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By


def get_data_from_url(url):
    driver = webdriver.Chrome()  # Replace with your preferred browser driver
    driver.get(url)

    # Wait for the table to load (adjust timeout as needed)
    wait_time = 10  # seconds
    driver.implicitly_wait(wait_time)

    # Find the table by ID
    table = driver.find_element(By.ID, "spectable")
    print(table.text)

    # Extract data from table rows within the "tbody" section
    rows = table.find_elements(By.TAG_NAME, "tr")  # Adjust based on actual element

    data = [[cell.text.strip() for cell in row.find_elements(By.TAG_NAME, "th")] for row in rows]

    driver.quit()  # Close the browser after data extraction

    return data


def save_data_to_excel(data, filename):
    wb = openpyxl.Workbook()
    sheet = wb.active

    # Add column headers if available in the first table row
    if data and data[0]:
        sheet.append(data[0])
        data = data[1:]  # Skip the header row

    for row in data:
        sheet.append(row)
    wb.save(filename)


if __name__ == "__main__":
    url = "https://hapiwangy.github.io/enterprise_visited/"
    data = get_data_from_url(url)
    print(data)

    if data:
        save_data_to_excel(data, "enterprise_visited.xlsx")
        print("Data saved successfully to enterprise_visited.xlsx")
    else:
        print("Error: Data extraction failed. No matching data found.")
