import pandas as pd
from selenium import webdriver
from bs4 import BeautifulSoup
from docx import Document

# Load the Word document containing URLs or paths to web pages
doc_path = 'data.docx'  # Update with your Word document path
doc = Document(doc_path)
urls = [p.text for p in doc.paragraphs if p.text.startswith('http')]  # Extract URLs

# Set up Chrome options to run the browser in incognito mode
chrome_options = webdriver.ChromeOptions()
chrome_options.add_argument("--incognito")

# Initialize the Chrome driver with the specified options
driver = webdriver.Chrome(options=chrome_options)

# Initialize lists to store extracted data
titles = []
descriptions = []
images = []
links = []

for url in urls:
    # Navigate to the URL from the Word document
    driver.get(url.strip())  # Remove any leading or trailing whitespace

    # Wait for the page to load
    driver.implicitly_wait(10)

    # Get the HTML content of the page after it has fully loaded
    html_content = driver.page_source

    # Parse the HTML content with BeautifulSoup
    soup = BeautifulSoup(html_content, 'html.parser')

    # Extract relevant information (modify this based on the structure of the web pages)
    title = soup.find('title').text.strip() if soup.find('title') else 'N/A'
    description = soup.find('meta', {'name': 'description'})['content'].strip() if soup.find('meta', {'name': 'description'}) else 'N/A'
    image = soup.find('img')['src'] if soup.find('img') else 'N/A'
    link = url.strip()  # Store the original URL as a link

    # Append the extracted data to the respective lists
    titles.append(title)
    descriptions.append(description)
    images.append(image)
    links.append(link)

# Close the Chrome driver
driver.quit()

# Create a Pandas DataFrame to store the extracted data
data = {
    'Title': titles,
    'Description': descriptions,
    'Image URL': images,
    'Link': links
}
df = pd.DataFrame(data)

# Save the DataFrame to an Excel file
excel_file_path = 'scraped_data.xlsx'  # Update with your desired Excel file path
df.to_excel(excel_file_path, index=False)

print("Data scraped and saved to Excel successfully!")
