# College Exam Results Scraper
# Author : Mostafa Ahmed Lotfy Moghazy
import requests
from bs4 import BeautifulSoup
import pandas as pd

# Initialize an empty list to store the scraped data
data = []

# Loop through the range of student codes
for code in range(29796, 30447):
    try:
        # Send a GET request to the target URL
        response = requests.get(
            f"http://app1.helwan.edu.eg/FaslBU/EngHelwan/HasasnUpMview.asp?StdCode={code}")

        # Skip iteration if the page is not found
        if response.status_code == 404:
            continue

        # Parse the page content
        soup = BeautifulSoup(response.content, "lxml")

        # Extract relevant data elements
        name_elements = soup.find_all("td", {'width': '360'})
        seat_number_elements = soup.find_all("td", {'width': '362'})
        results_elements = soup.find_all("td", {'width': '100'})
        c2_results_elements = soup.find_all("td", {'width': '81'})

        def extract_text(elements, font_face='Traditional Arabic'):
            """
            Extracts text from a specific font face within the given elements.
            """
            if elements:
                font_element = elements[0].find('font', {'face': font_face})
                if font_element:
                    return font_element.text.strip()
            return ""

        def extract_bold_text(elements, index):
            """
            Extracts bold text from the specified index of the elements.
            """
            try:
                bold_element = elements[index].find('b')
                if bold_element:
                    return bold_element.text.strip()
            except (IndexError, AttributeError):
                return ""
            return ""

        # Extract the student name and seat number
        student_name = extract_text(name_elements)
        seat_number = extract_bold_text(seat_number_elements, 0)

        # Proceed only if seat number exists
        if seat_number:
            # Extract the main results and C2 results
            main_results = [extract_bold_text(
                results_elements, i) for i in range(11)]
            c2_results = [extract_bold_text(
                c2_results_elements, i) for i in range(2)]

            # Append the data to the list
            data.append([student_name, seat_number] +
                        main_results + c2_results)

    except Exception as e:
        print(f"An error occurred for code {code}: {e}")

# Define the column names for the DataFrame
columns = [
    "Name", "Seat Number", "Total", "Math1", "Phy1", "Mech",
    "Chem", "Cs", "Math2", "Phy2", "ED", "Prod", "E", "tot", "HR"
]

# Create a DataFrame and save it to an Excel file
df = pd.DataFrame(data, columns=columns)
output_file = "C:\\Users\\Mostafa\\Desktop\\Helwan_Final.xlsx"
df.to_excel(output_file, index=False)

print(f"Data successfully saved to {output_file}")
