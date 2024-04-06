# Excelbnb: Airbnb Listings to Excel
##### Excelbnb is a Python application designed to convert Airbnb listings into a structured Excel spreadsheet. This tool is perfect for users looking to analyze Airbnb listings data for specific locations without manually sorting through the website. It fetches listing details such as the name, type, price per night, rating, and listing URL. It then organizes this information into an easily readable Excel file.

## Features
- Fetch Airbnb listings based on URL.<br>
- Sort listings by price per night.<br>
- Output data includes listing name, type, price, rating, and URL.<br>
- Data is saved into an Excel spreadsheet with formatting for better readability.

## Installation

### Dependencies
- This application requires Python 3.x and Google Chrome<br>
- You can install the required packages using pip:<br>
`pip install requests pandas bs4 selenium webdriver_manager openpyxl lxml`

### Setting Up
- Clone the repository to your local machine. <br>


### Usage
- Before running the application, you need to set the URL of the Airbnb location you are interested in. <br>
- Open the script and navigate to the CHANGE AIRBNB URL HERE section:<br>
`# ---!!!CHANGE AIRBNB URL HERE!!!---`<br>
`url = 'your Airbnb link here'`<br>
- Replace 'your Airbnb link here' with the URL of the Airbnb search result you want to analyze.<br>
- To run the application, execute the script in your terminal or command prompt:<br>
`python Excelbnb.py`<br>
- Wait for the automated Chrome browser window to close automatically after fetching the data. Once the script completes its execution, you'll find the Airbnb.xlsx file in your directory, containing the structured data.

## Contributing
Contributions to Excelbnb are welcome! Feel free to fork the repository, make your changes, and submit a pull request.

### License
[MIT License](https://github.com/TinsleyDevers/Excelbnb/blob/main/LICENSE)
