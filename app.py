import openpyxl
import requests
from bs4 import BeautifulSoup


def get_data_from_url(url):
    response = requests.get(url)
    soup = BeautifulSoup(response.content, "html.parser")

    # Find the specific table with ID "spectable"
    table = soup.find("table", id="spectable")

    if not table:
        print("Error: Table with ID 'spectable' not found.")
        return None

    # Extract data from table rows within the "thead" section
    rows = table.find("thead").find_all("tr")
    data = [[cell.text.strip() for cell in row.find_all("td")] for row in rows]

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

    if data:
        save_data_to_excel(data, "enterprise_visited.xlsx")
        print("Data saved successfully to enterprise_visited.xlsx")
    else:
        print("Error: Data extraction failed.")
