# Karnataka High Court Automation Script

This script automates the process of downloading judgment PDFs from the Karnataka High Court website. It uses Selenium for browser automation, OpenPyXL for Excel file manipulation, and Google Gemini for CAPTCHA solving.

## Features
- Automates the selection of bench, date range, and CAPTCHA solving.
- Downloads judgment PDFs and renames them based on case details.
- Updates an Excel file with case details and download status.
- Handles dynamic table rows and blocking elements on the website.
- Automatically retries failed downloads and updates configuration files.

## Prerequisites
1. **Python**: Ensure Python 3.8 or higher is installed.
2. **Google Chrome**: Install Google Chrome on your system.
3. **ChromeDriver**: Download the ChromeDriver version matching your Chrome browser from [here](https://googlechromelauncher.github.io/chromedriver/). Place the `chromedriver.exe` in the same directory as the script.
4. **Python Libraries**: Install the required Python libraries using the following command:
   ```bash
   pip install -r requirements.txt
   ```
   The required libraries include:
   - `selenium`
   - `openpyxl`
   - `requests`
   - `python-dotenv`
   - `google-generativeai`

5. **Environment Variables**: Create a `.env` file in the script directory with the following content:
   ```
   GEMINI_API_KEY=<Your_Google_Gemini_API_Key>
   ```

6. **Configuration Files**:
   - `test.json`: Contains configuration details such as date range, download directory, and Excel file path.
   - `status.xlsx`: Tracks the status of tasks (e.g., pending, processing, completed).

## How It Works
1. **Configuration**:
   - The script reads the `test.json` file for configuration details, including the date range, download directory, and Excel file path.
   - If `downloaded_pdf_number` is `-1`, the script updates the `end_serial` value by navigating to the website and retrieving the last serial number.

2. **Browser Setup**:
   - The script sets up a Selenium WebDriver with Chrome options to handle downloads and bypass browser security warnings.

3. **CAPTCHA Solving**:
   - The script uses Google Gemini to solve CAPTCHA images automatically. If CAPTCHA solving fails, manual intervention is required.

4. **PDF Download**:
   - The script navigates to the Karnataka High Court website, selects the bench, enters the date range, and searches for judgments.
   - It iterates through the table rows, downloads PDFs, and renames them based on case details (e.g., case number, year, parties involved).

5. **Excel Updates**:
   - The script updates an Excel file with case details, including the status of the PDF (e.g., downloaded, not available).

6. **Error Handling**:
   - The script handles errors such as missing files, CAPTCHA failures, and website issues. It retries failed downloads and logs errors for debugging.

7. **Task Completion**:
   - Once all PDFs are downloaded, the script marks the task as completed in `status.xlsx` and resets the configuration for the next batch.

## Usage
1. **Run the Script**:
   - To start the script, run the following command:
     ```bash
     python test.py
     ```

2. **Update Serial Numbers**:
   - If `downloaded_pdf_number` is `-1`, run the script in update-serial mode:
     ```bash
     python test.py update-serial
     ```

3. **Manual CAPTCHA Solving**:
   - If the script fails to solve the CAPTCHA, solve it manually in the browser and press Enter to continue.

## File Structure
- `test.py`: Main script for automation.
- `test.json`: Configuration file for date range, download directory, and Excel path.
- `status.xlsx`: Tracks task status and total PDFs.
- `.env`: Contains the Google Gemini API key.

## Logs and Debugging
- The script logs progress and errors to the console.
- Check the console output for troubleshooting steps if the script encounters issues.

## Notes
- Ensure the ChromeDriver version matches your installed Chrome browser.
- The script assumes the website structure remains consistent. If the website changes, the XPath selectors may need to be updated.
- Use a valid Google Gemini API key for CAPTCHA solving.

## Disclaimer
This script is intended for educational and personal use only. Ensure compliance with the terms and conditions of the Karnataka High Court website before using this script.
