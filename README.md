# Ecwid Sales Data Excel Formatter

This project is a Python script that retrieves sales data from the Ecwid API and formats it into an Excel file (`SalesData.xlsx`). The Excel file contains detailed sales information, including images, customization details, and customer information. It alternates the row background color and border styles for each customer to enhance visual clarity.

## Features

- Fetches sales data from the Ecwid API using an API token and store ID.
- Organizes sales information such as order number, image URL, item size, customization, and customer details.
- Applies alternating background colors (blue and green) to differentiate each customer’s records.
- Adds thick black borders around each customer’s set of rows for better separation.
- Adjusts column widths and row heights dynamically based on content.
- Supports image display directly within the Excel sheet using the `IMAGE()` function.

## Prerequisites

Make sure you have Python installed, and install the following dependencies using `pip`:

```bash
pip install openpyxl requests
