# Document Metadata Extract

## Project description

The challenge was to identify which of several hundred documents on a public website had incomplete metadata fields.

The solution was to automate the process by developing an application in Python to read in the dataset as a csv, iterate over website URLs of PDF and Excel documents, open the documents, extract metadata and output information to a spreadsheet to be able to identify missing fields.

Populating document fields such as Author, Title (equivalent to a web page title tag) and Subject (equivalent to a web page meta description) assists with Search Engine Optimisation by providing metadata for search engines to crawl and create listings for documents, and to determine their search rankings.

## Data source

Test dataset created for the purpose of testing the algorithm. 

## Requirements

* Python 3.8.x
* requests: a Python HTTP library 
* PyPDF2: a Python library built as a PDF toolkit
* openpyxl: a Python library to read/write Excel files