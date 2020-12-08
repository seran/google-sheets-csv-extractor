# Google Sheets customisable CSV exporter

The application uses the Google Sheets Python API library to export custom CSV
from sheet. Refer the below links for more details:

- https://developers.google.com/sheets/api/guides/authorizing
- https://developers.google.com/sheets/api/quickstart/python

Enable the Sheets API and move the credentials.json file to the program
directory.

Steps to execute:


1. After the successful login, run the meta command to generate the metadata file.
2. You can use the enabled field to activate/deactivate the necessary fields to
export.
3. Then run the extract command to pull the data from the sheets.