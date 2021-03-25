RPA Challenge - IT Dashboard
https://www.notion.so/RPA-Challenge-IT-Dashboard-ec59bc2659e64323a7af99fcd4d24c21

Extracting data from itdashboard.gov:
1. 
- Get a list of agencies and the amount of spending from the main page
- Write the amounts to an Excel file and call the sheet "Agencies"

2.
- Select one of the agencies, for example, National Science Foundation (this should be configured in a file or on a Robocloud)
- Going to the agency page scrape a table with all "Individual Investments" and write it to a new sheet in Excel.
- If the "UII" column contains a link, open it and download PDF with Business Case (button "Download Business Case PDF")
  (You need to get the data from Section A in each PDF. Then compare the value "Name of this Investment" with the column 
   "Investment Title", and the value "Unique Investment Identifier (UII)" with the column "UII")
3.   
- Store downloaded files and Excel sheet to the root of the output folder.
- Use RPA Framework: https://rpaframework.org/ ( pip install rpaframework )

How to run:
Place the list of the agencies to obtain the detailed information into 'config.txt'
Start the virtual environment for the project
From the project folder enter: python bot.py
Open saved results from project output folder
