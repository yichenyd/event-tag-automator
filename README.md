# EventBadge-Automator

A Python-based tool that automates the generation of printable name tags from Excel registration rosters. It handles data cleaning, dynamic room assignment mapping, and PDF layout generation with conditional visual indicators.

## Features
* **Data Extraction & Cleaning**: Parses raw Excel sheets to filter valid participants and standardize sport categories (e.g., Soccer, Dance, Badminton, Boxing).
* **Dynamic Room Mapping**: Searches and links up to two room assignments per participant based on cross-sheet data.
* **Automated PDF Generation**: Renders an 8x2 grid of name tags on letter-sized pages.
* **Visual Indicators**: Conditionally overlays a star icon for participants with specific food requirements.

## File Structure
* `DataSheet.py`: The data processing script. It reads the raw Excel registration file, filters and formats the fields, maps the room numbers, and outputs a clean data sheet.
* `namePDF.py`: The PDF generation script. It reads the cleaned data sheet and uses ReportLab to draw the name tags, text, and conditional visual overlays onto the final PDF.
* `star.png`: The icon used to indicate dietary restrictions.

## Usage
1. Ensure your raw data file is properly formatted and the file paths in `DataSheet.py` point to your source Excel file.
2. Run the data processing script to generate the cleaned dataset:

       python DataSheet.py

3. Run the PDF generation script to create the tags:

       python namePDF.py

4. Open the generated `NamePDF.pdf` to view and print your ready-to-use name tags.

