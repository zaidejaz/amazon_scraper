# Amazon Scraper

This script scrapes Amazon product reviews and updates an Excel file with the latest review data.

## Prerequisites

- Python 3.10 or higher
- Poetry (for package management)
- Google Chrome browser

## Setup

1. Install Poetry if you haven't already:
   ```
   pip install poetry
   ```

2. Navigate to the project directory in your terminal.

3. Install the required dependencies using Poetry:
   ```
   poetry install
   ```

## Usage

1. Place your Excel file (named `data.xlsx`) in the same directory as the script.

2. Run the script using Poetry:
   ```
   poetry run python amazon_scraper.py
   ```

3. The script will process each Amazon URL in the Excel file, scrape the review data, and update the file with the new information.

4. Once completed, you'll find the updated data in `data.xlsx`.

## Notes

- The script uses undetected_chromedriver to avoid detection. Make sure you have Google Chrome installed on your system.
- If you encounter a captcha, the script will attempt to solve it automatically.
- The script includes a few-second delay between requests to avoid overwhelming Amazon's servers.

## Troubleshooting

If you encounter any issues:

1. Make sure all dependencies are correctly installed.
2. Check that your `data.xlsx` file is in the correct format and location.
3. Ensure you have a stable internet connection.
4. If captchas are failing wait before trying again.

For any persistent problems, please contact the developer.