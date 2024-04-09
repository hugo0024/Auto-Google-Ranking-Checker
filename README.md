# Auto Google Ranking Checker

This project is a Python application that checks the Google search ranking of specific keywords for different companies.

## Features

- Add and remove target keywords
- Add and remove target sites
- Display the keyword, rank, and page in separate text areas
- Save the search results in an Excel file

## Usage

1. Run `rank.py` to start the application.
2. Select a company from the dropdown list. The application will automatically load the keywords and target site URLs for the selected company.
3. Use the "Add Keyword" and "Remove Keyword" buttons to manage the keywords.
4. Use the "Add Site" and "Remove Site" buttons to manage the target sites.
5. Click the "Search" button to start the search. The application will display the keyword, rank, and page in separate text areas.

## Dependencies

- tkinter
- requests
- BeautifulSoup
- openpyxl
- pandas

## Note

The application uses desktop user agent strings to simulate a desktop browser when sending requests to Google.

## License

This project is licensed under the terms of the MIT license.