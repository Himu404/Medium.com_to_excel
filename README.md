# Medium Articles Scraper

This Python script scrapes titles and links of Medium articles related to AWS from the Medium website and stores them in an Excel file.

## Features
- Uses Selenium WebDriver to interact with Medium's website.
- Uses BeautifulSoup for HTML parsing.
- Extracts up to 100 article titles and links.
- Stores the extracted data in an Excel file.

## Prerequisites
- Python 3.x
- Google Chrome
- ChromeDriver
- Required Python packages:
  - selenium
  - beautifulsoup4
  - openpyxl

## Installation

1. **Clone the repository:**

    ```bash
    git clone https://github.com/your-username/medium-articles-scraper.git
    cd medium-articles-scraper
    ```

2. **Install the required packages:**

    ```bash
    pip install selenium beautifulsoup4 openpyxl
    ```

3. **Download ChromeDriver:**
    - Ensure it matches your Chrome browser version.
    - [Download ChromeDriver](https://sites.google.com/a/chromium.org/chromedriver/downloads)
    - Place the `chromedriver.exe` file in a known directory (update the path in the script if necessary).

## Usage

1. **Update the script:**
    - Ensure the `webdriver_path` variable in the script points to the location of your `chromedriver.exe`.

    ```python
    webdriver_path = r'C:\path\to\your\chromedriver.exe'
    ```

2. **Run the script:**

    ```bash
    python scraper.py
    ```

3. **Result:**
    - An Excel file named `medium_articles.xlsx` will be created in the same directory, containing the titles and links of the articles.

## Script Explanation

### fetch_titles_and_links(url, limit=100)
- This function uses Selenium to open the given URL, scrolls through the page to load content, and uses BeautifulSoup to parse and extract up to 100 article titles and links.

### write_to_excel(titles_and_links)
- This function takes the list of titles and links and writes them to an Excel file named `medium_articles.xlsx`.
