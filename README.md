# Powershell-Scripts
A repository for Powershell  scripts written to automate repetitive tasks.

Several scripts have variables containing "Removed" string value.  These were removed for confidentiality purposes.

These scripts were written very specifically for very repetitive and long tasks given to me while I held a temporary position.
They are purely for demonstration and are not expected to function for other users.


FileName_Update:
This script was written because of a policy that each file name needed to reflect the date that it was last modified.  Because there were over 400 files, I wrote this script to update the file name of all documents to append the current date (and remove the old date) so that I didn't need to rename each file individually.

LinkRemover:
Originally I had been tasked with opening over 400 word documents and removing a particular line from each document.  Rather than do it manually, I wrote the LinkRemover Powershell script to open all 400+ documents and remove the line.

OfficeLocation_Finder:
The task was to take a given excel spreadsheet with a column of 300+ employee names and fill in the second column with their office location.  I did this using Powershell, Excel, and Outlook directory.

Textbook_Discrepancy_Checker:
Written to check discrepancies in information regarding hundreds of textbooks, i.e, version number, publishers, authors, etc.
