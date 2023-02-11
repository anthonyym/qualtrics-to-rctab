# Qualtrics to RCTab

Qualtrics to RCTab is an application that takes in a Qualtrics CSV file containing the ballots for a ranked choice voting (RCV) race, or multiple races. For every race within the Qualtrics CSV file, the application creates (1) an ES&S CVR Excel file and (2) a configuration JSON file with a single-winner RCV configuration, and then calls [RCTab](https://www.rcvresources.org/rctab), a ranked choice voting tabulator, with these files as input.
