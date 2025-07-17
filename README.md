# StudentProgramsReports

Code to automate billing PDF generation for student programs and districts.

## Overview

This repository contains scripts primarily written in Google Apps Script (JavaScript-based) to assist with:
- Automating the generation of billing PDFs for student programs.
- Organizing Google Drive folders by school/program.
- Processing and transforming spreadsheet data for reporting.

## Features

- **Automated PDF Generation**: Export filtered spreadsheet data as PDFs for each school/program and save to designated folders in Google Drive.
- **Google Drive Folder Management**: Automatically create folders for each school, program, or district.
- **Spreadsheet Data Processing**: Extract, filter, and organize student/program data for reporting and billing.
- **Flexible Configuration**: Easily configure tab names, ranges, and columns to keep via the `nsSettings.gs` file.

## File Structure

- `nsSettings.gs`: Central configuration for tab names, ranges, and folder IDs.
- `nsFolder.gs`: Functions for creating folders in Google Drive based on spreadsheet data.
- `nsHelpers.gs`: Utility functions for column extraction, date formatting, folder listing, and 2D array conversion.
- `pdfHelpers.gs`: Functions for exporting spreadsheet ranges to PDF and saving them to specific folders.
- `runReports.gs`: Main logic for running reports and exporting PDFs using configured filters.
- `practice.gs`: Example and utility functions for manipulating and updating spreadsheet data.

## Setup

1. **Google Apps Script Project**: Deploy the scripts within a Google Apps Script project attached to your billing/reporting spreadsheet.
2. **Configuration**:
   - Update `nsSettings.saveFolderID` to your target Google Drive folder.
   - Adjust tab names and ranges in `nsSettings.gs` to match your spreadsheet structure.
3. **Permissions**:
   - Ensure the script has permission to access Google Drive and Google Sheets.
   - All users running the script must have edit access to the destination folder.

## Usage

- **Create Folders**: Run the function `runCreateFolders` to create folders for each school/program listed in your spreadsheet.
- **Generate Billing PDFs**: Run the `runReports` function to process the spreadsheet and generate PDFs for each school/program, saving them to the appropriate folder.
- **Update Student Info**: Use functions in `practice.gs` to extract and update student information across tabs.

## Contributing

If you'd like to contribute:
1. Fork the repository.
2. Submit a pull request with a detailed description of your changes.
3. Ensure your code follows the existing style and is well-documented.

## License

This repository currently does not specify a license. Please contact the repository owner for usage or redistribution permissions.

## Maintainer

- [oshag-mcesc](https://github.com/oshag-mcesc)

---

*This README was generated based on repository code and configuration. For additional documentation or help, please open an issue in this repository.*
