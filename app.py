import concurrent.futures

import openpyxl
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support import expected_conditions as ec
from selenium.webdriver.support.ui import WebDriverWait


def driver_get_text(driver, locator):
    """Get text of element.

    :param driver: WebDriver instance.
    :param locator: Locator of element.
    :return: Text of element.
    """
    return WebDriverWait(driver, 10).until(ec.presence_of_element_located(locator)).text


def get_data_from_url(url):
    driver = webdriver.Chrome()  # Replace with your preferred browser driver
    driver.get(url)

    # Wait for the table to load (adjust timeout as needed)
    wait_time = 10  # seconds
    driver.implicitly_wait(wait_time)

    data = []
    for i in range(1, 11):  # Get column headers
        data.append(driver_get_text(driver, (
            By.XPATH, f"/html/body/div/div/div/table/thead/tr[2]/td[{i}]")))

    # Get table data
    with concurrent.futures.ThreadPoolExecutor() as executor:
        rows_data = list(
            executor.map(lambda row_num: get_data_from_row(driver, row_num), range(2, 5002)))
    data.extend(rows_data)

    driver.quit()  # Close the browser after data extraction
    print(data)
    return data


def get_data_from_row(driver, row_num):
    temp = []
    for j in range(1, 11):
        temp.append(driver_get_text(driver, (
            By.XPATH, f"/html/body/div/div/div/table/thead/tr[{row_num}]/td[{j}]")))
    print(f"{row_num} done")
    return temp


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

    if data:
        save_data_to_excel(data, "enterprise_visited.xlsx")
        print("Data saved successfully to enterprise_visited.xlsx")
    else:
        print("Error: Data extraction failed. No matching data found.")
