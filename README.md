# Send a mail with Excel

## Description

This script automates the process of extracting data from Excel files, merging it with recipient addresses, and sending personalized emails with the data presented in an HTML table format. It is designed to streamline data distribution and reporting workflows.

## Dependencies

*   `pandas`: For data manipulation and reading Excel files.
*   `openpyxl`: For more advanced Excel operations, such as reading hyperlinks.
*   `win32com.client`: For interacting with Microsoft Outlook to send emails (Windows only).
*   `pathlib`: For handling file paths.
*   `re`: For regular expressions, used in parsing hyperlink formulas.
*   `glob`: For finding files using patterns.

To install the dependencies, run:

```bash
pip install pandas openpyxl pywin32
```

## Usage

1.  **Clone the repository:**

    ```bash
    git clone <repository_url>
    cd <repository_directory>
    ```

2.  **Configure the script:**

    *   Modify the `main_multi_excel` function in `Send a mail with Excel.ipynb` to set the following variables:
        *   `EXCEL_FOLDER`: The path to the folder containing the Excel files.
        *   `ADDRESS_BOOK_PATH`: The path to the Excel file containing the recipient addresses (with columns "AddressName" and "Name").
        *   `EMAIL_SUBJECT`: The subject of the email.
        *   `MESSAGE`: The body of the email.

3.  **Run the script:**

    Open the `Send a mail with Excel.ipynb` Jupyter Notebook and run all cells.

## Configuration

The following configurations are required in the `main_multi_excel` function:

*   `EXCEL_FOLDER`:  The folder path where your Excel files are located.
*   `ADDRESS_BOOK_PATH`: The path to the Excel file containing the recipient addresses. This file should have two columns: "Name" and "AddressName".
*   `EMAIL_SUBJECT`: The subject line for the emails that will be sent.
*   `MESSAGE`: The main body text of the email.

## Address Book Format

The address book should be an Excel file with the following columns:

*   `Name`: The name of the recipient (used to match entries in the Excel data files).
*   `AddressName`: The email address of the recipient.

## License

[Specify the license here, e.g., MIT License]
