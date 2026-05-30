# WooCommerce Order Fetcher

Google Apps Script to fetch processing orders from a WooCommerce REST API and populate a Google Sheet.

## Features
* Clears the target sheet before importing.
* Imports up to 100 processing orders starting at row 10.
* Dynamically parses order metadata, billing info, and custom field data.
* Automatically hides remaining empty rows to keep the spreadsheet clean.
