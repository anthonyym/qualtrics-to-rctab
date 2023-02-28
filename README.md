# Qualtrics to RCTab

Qualtrics to RCTab is a Windows application that takes in a Qualtrics CSV file containing the ballots for a single-winner ranked choice voting (RCV) race, or multiple races. For every race within the Qualtrics CSV file, the application creates (1) an ES&S CVR Excel file and (2) a configuration JSON file with a single-winner RCV configuration, and then calls [RCTab](https://www.rcvresources.org/rctab), a ranked choice voting tabulator, with these files as input.

## Installation and Usage
The following link contains a precomiled version of the application, sample input files, and a Readme for how to install and run the application: https://drive.google.com/drive/folders/1TGBoAOR2aNhcy1jTCQl3t46IAA9pqXUu?usp=share_link

## Developer Setup
This software was tested and run with:
 - Windows 11
 - [Python 3.6.3](https://www.python.org/downloads/release/python-363/)
 - Package dependencies to install (for example via `pip install [package_name]==[package_version]`):
   - wxPython==4.1.1
   - pandas==1.1.5
